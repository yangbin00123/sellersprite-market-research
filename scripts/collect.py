#!/usr/bin/env python3
"""卖家精灵市场大盘调研：给一个关键词，自动下载 15 张表。

用法:  python collect.py "nose rings" [--out ~/SellerSprite调研] [--site 美国站] [--step all|products|months|keywords|conv]
首次运行会打开浏览器，请在窗口里登录卖家精灵（登录态保存在 ~/.sellersprite-profile，之后免登录）。
已下载的文件会跳过，中途失败可直接重跑。失败时在输出目录生成 debug_*.png 供排查。
"""
import argparse, json, re, sys, time
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---- 页面地址（卖家精灵改版时只需改这里）----
URL_COMPETITOR = "https://www.sellersprite.com/v3/competitor-lookup"
URL_EXPAND = "https://www.sellersprite.com/v3/expand-keywords"
URL_CONV = "https://www.sellersprite.com/v3/keyword-conversion"
TOP_ASIN = 20          # 取月销量前 N 个商品做关键词反查
TRAFFIC_CUM = 0.95     # 累计流量占比阈值
MAX_KW = 1000          # 转化率工具单次上限
N_MONTHS = 12


def log(*a): print("[collect]", *a, flush=True)


class Bot:
    def __init__(self, page, out: Path, logs: Path, site: str):
        self.p, self.out, self.logs, self.site = page, out, logs, site
        self.mf = logs / "manifest.json"          # 记录 步骤标签 -> 下载文件名（不放进结果文件夹）
        self.done_map = json.loads(self.mf.read_text("utf8")) if self.mf.exists() else {}

    def done(self, tag):
        f = self.done_map.get(tag)
        return bool(f) and (self.out / f).exists()

    # ---------- 通用 ----------
    def shot(self, name):
        try: self.p.screenshot(path=str(self.logs / f"debug_{name}.png"), full_page=True)
        except Exception: pass

    def ensure_login(self):
        for _ in range(300):  # 最多等 5 分钟
            if "login" not in self.p.url and "sign" not in self.p.url and \
               self.p.get_by_text(re.compile("登录|登陆")).count() == 0:
                return
            if _ == 0: log("请在浏览器窗口登录卖家精灵（登录后自动继续）…")
            time.sleep(1)
        sys.exit("等待登录超时")

    def goto(self, url):
        self.p.goto(url, wait_until="domcontentloaded")
        self.p.wait_for_timeout(2500)
        self.ensure_login()
        if self.p.url.split("?")[0] != url:
            self.p.goto(url, wait_until="domcontentloaded"); self.p.wait_for_timeout(2500)

    def click_text(self, text, exact=True, timeout=8000):
        loc = self.p.get_by_text(text, exact=exact).locator("visible=true").first
        loc.click(timeout=timeout)
        self.p.wait_for_timeout(400)

    def set_site(self):
        if self.site and self.p.get_by_text(self.site, exact=True).locator("visible=true").count() == 0:
            self.click_text("美国站", exact=False); self.click_text(self.site)

    def open_period(self):
        """点开时间下拉（当前显示 最近30天 / 2026.08 / 近90天 之类）。"""
        cur = self.p.get_by_text(re.compile(r"^(最近30天|近90天|\d{4}[.\-]\d{2})$")).locator("visible=true").first
        cur.click(); self.p.wait_for_timeout(500)

    def period_options(self):
        items = self.p.locator("li:visible, .el-select-dropdown__item:visible, [role=option]:visible").all_inner_texts()
        return [t.strip() for t in items if t.strip()]

    def pick_period(self, label):
        self.open_period()
        self.p.locator("li:visible, .el-select-dropdown__item:visible, [role=option]:visible").filter(
            has_text=re.compile(rf"^\s*{re.escape(label)}\s*$")).first.click()
        self.p.wait_for_timeout(500)

    def query(self):
        self.p.get_by_text(re.compile(r"^\s*立即查询\s*$")).locator("visible=true").first.click()
        self.p.wait_for_load_state("networkidle", timeout=60000)
        self.p.wait_for_timeout(2500)

    def export(self, tag):
        """点 导出，保持卖家精灵原文件名保存到结果文件夹；若弹出确认框则点确认。"""
        if self.done(tag): log("已存在，跳过", self.done_map[tag]); return
        btn = self.p.get_by_role("button", name=re.compile(r"^\s*导出\s*$")).locator("visible=true")
        if btn.count() == 0:
            btn = self.p.get_by_text(re.compile(r"^\s*导出\s*$")).locator("visible=true")
        try:
            with self.p.expect_download(timeout=120000) as d:
                btn.first.click()
                self.p.wait_for_timeout(1500)
                for t in ("确认导出", "确认", "确定", "导出"):  # 可能的二次确认
                    c = self.p.locator(".el-dialog:visible, .el-message-box:visible, .modal:visible, [role=dialog]:visible").get_by_text(t, exact=True)
                    if c.count(): c.first.click(); break
            name = d.value.suggested_filename
            d.value.save_as(self.out / name)
            self.done_map[tag] = name; self.mf.write_text(json.dumps(self.done_map, ensure_ascii=False, indent=1), "utf8")
            log("已保存", name)
        except PWTimeout:
            self.shot(tag); sys.exit(f"导出超时：{tag}（截图见 {self.logs}；可能是当日导出次数用完或页面改版）")

    # ---------- 1/2 竞品表 ----------
    def products(self, kw):
        if self.done("products"): return
        self.goto(URL_COMPETITOR); self.set_site()
        self.p.get_by_placeholder(re.compile("关键词")).first.fill(kw)
        self.query()
        # 不展示变体 + 月销量降序
        cb = self.p.get_by_text("展示所有变体").locator("visible=true").first
        if self.p.locator("label.is-checked:has-text('展示所有变体')").count(): cb.click()
        self.sort("月销量", "降序")
        self.export("products")

    def sort(self, field, order):
        for val in (field, order):
            sel = self.p.locator(".el-select:visible").filter(has_text=re.compile(r"月销量|流量占比|降序|升序|搜索量|购买量"))
            # 逐个下拉按当前显示文本匹配
            for i in range(sel.count()):
                if val in ("月销量", "流量占比") and re.search("月销量|流量占比|搜索量|购买量", sel.nth(i).inner_text()):
                    sel.nth(i).click(); self.pick_option(val); break
                if val in ("降序", "升序") and re.search("降序|升序", sel.nth(i).inner_text()):
                    sel.nth(i).click(); self.pick_option(val); break
        self.p.get_by_text("确定", exact=True).locator("visible=true").first.click()
        self.p.wait_for_load_state("networkidle", timeout=60000); self.p.wait_for_timeout(2000)

    def pick_option(self, text):
        self.p.locator("li:visible, .el-select-dropdown__item:visible, [role=option]:visible").filter(
            has_text=re.compile(rf"^\s*{text}\s*$")).first.click()
        self.p.wait_for_timeout(300)

    def months(self, kw):
        self.goto(URL_COMPETITOR); self.set_site()
        self.p.get_by_placeholder(re.compile("关键词")).first.fill(kw)
        self.open_period()
        opts = [o for o in self.period_options() if re.fullmatch(r"\d{4}[.\-]\d{2}", o)]
        self.p.keyboard.press("Escape")
        months = sorted(opts, reverse=True)[:N_MONTHS]
        if len(months) < N_MONTHS: log(f"警告：下拉里只有 {len(months)} 个月份：{months}")
        for m in sorted(months):
            tag = f"month:{m.replace('-', '.')}"
            if self.done(tag): continue
            self.pick_period(m); self.query()
            if self.p.locator("label.is-checked:has-text('展示所有变体')").count():
                self.p.get_by_text("展示所有变体").locator("visible=true").first.click()
            self.sort("月销量", "降序"); self.export(tag)

    # ---------- 3 反查 ----------
    def asins(self):
        wb = openpyxl.load_workbook(self.out / self.done_map["products"], read_only=True)
        rows = wb.active.iter_rows(values_only=True)
        head = next(rows); i = head.index("ASIN")
        out = []
        for r in rows:
            if r[i] and r[i] not in out: out.append(str(r[i]).strip())
            if len(out) >= TOP_ASIN: break
        return out

    def keywords_table(self):
        if self.done("expand"): return
        asins = self.asins(); log("反查 ASIN：", ",".join(asins))
        self.goto(URL_EXPAND); self.set_site()
        box = self.p.locator("input:visible, textarea:visible").first
        box.click(); self.p.keyboard.insert_text(",".join(asins)); self.p.keyboard.press("Enter")
        self.query()
        self.sort("流量占比", "降序")
        self.export("expand")

    # ---------- 4 转化率 ----------
    def kw_list(self):
        wb = openpyxl.load_workbook(self.out / self.done_map["expand"], read_only=True)
        rows = wb.worksheets[0].iter_rows(values_only=True)
        head = next(rows); ik, ip = head.index("关键词"), head.index("流量占比")
        cum, kws = 0.0, []
        for r in rows:
            v = r[ip]
            v = float(v.strip("%")) / 100 if isinstance(v, str) else float(v or 0)
            kws.append(str(r[ik]).strip()); cum += v
            if cum >= TRAFFIC_CUM or len(kws) >= MAX_KW: break
        log(f"累计流量占比 {cum:.2%}，共 {len(kws)} 个关键词")
        return kws

    def conversion(self):
        if self.done("conv"): return
        kws = self.kw_list()
        (self.logs / "核心关键词清单.txt").write_text("\n".join(kws), encoding="utf8")
        self.goto(URL_CONV); self.set_site()
        self.pick_period("近90天")
        box = self.p.locator("textarea:visible, input:visible").first
        box.click(); self.p.keyboard.insert_text("\n".join(kws))
        self.query(); self.export("conv")

    def check(self):
        total = 2 + N_MONTHS + 1
        n = sum(self.done(t) for t in self.done_map)
        (self.logs / "采集清单.md").write_text("\n".join(f"- {t}: {f}" for t, f in self.done_map.items()), encoding="utf8")
        log(f"完成：{n}/{total} 张", "" if n == total else "← 不完整，请重跑补齐")


def recon(b):
    """保存每个页面的截图 + 可见按钮/输入框/下拉清单，页面改版时发给开发者即可适配。"""
    for name, url in (("competitor", URL_COMPETITOR), ("expand", URL_EXPAND), ("conv", URL_CONV)):
        b.goto(url); b.p.wait_for_timeout(3000)
        b.shot(f"recon_{name}")
        info = b.p.evaluate("""() => [...document.querySelectorAll('button,input,textarea,select,a,[role=button],[role=tab],.el-select,label')]
          .filter(e => e.offsetParent).slice(0, 300)
          .map(e => [e.tagName, e.className.toString().slice(0, 60), (e.innerText || e.placeholder || '').trim().slice(0, 30)].join(' | '))""")
        (b.logs / f"recon_{name}.txt").write_text(b.p.url + "\n" + "\n".join(info), encoding="utf8")
    log("侦察完成，请把", b.logs, "里的 recon_* 文件发给开发者")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("keyword")
    ap.add_argument("--out", default="~/SellerSprite调研")
    ap.add_argument("--site", default="美国站")
    ap.add_argument("--cdp", help="连接你已登录的 Chrome，如 http://localhost:9222（Chrome 需以 --remote-debugging-port=9222 启动）")
    ap.add_argument("--recon", action="store_true", help="侦察模式：只保存各页面截图和控件清单到 _logs，用于适配改版")
    ap.add_argument("--step", default="all", choices=["all", "products", "months", "keywords", "conv"])
    a = ap.parse_args()
    out = Path(a.out).expanduser() / re.sub(r'[\\/:*?"<>|]', "_", a.keyword.strip())
    out.mkdir(parents=True, exist_ok=True)
    logs = Path(a.out).expanduser() / "_logs" / out.name   # 调试截图/状态记录放在结果文件夹之外
    logs.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as pw:
        if a.cdp:
            ctx = pw.chromium.connect_over_cdp(a.cdp).contexts[0]
            page = ctx.new_page()
        else:
            ctx = pw.chromium.launch_persistent_context(
                str(Path("~/.sellersprite-profile").expanduser()), headless=False,
                accept_downloads=True, viewport={"width": 1440, "height": 900}, locale="zh-CN")
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.set_default_timeout(15000)
        b = Bot(page, out, logs, a.site)
        if a.recon: return recon(b)
        try:
            if a.step in ("all", "products"): b.products(a.keyword)
            if a.step in ("all", "months"): b.months(a.keyword)
            if a.step in ("all", "keywords"): b.keywords_table()
            if a.step in ("all", "conv"): b.conversion()
            b.check()
        except Exception as e:
            b.shot("error"); raise
        finally:
            if not a.cdp: ctx.close()


if __name__ == "__main__":
    main()
