#!/usr/bin/env python3
"""卖家精灵市场大盘调研：给一个关键词，自动下载 15 张表。

用法:  python collect.py "nose rings" [--out ~/SellerSprite调研] [--site 美国站] [--step all|products|months|keywords|conv]
首次运行会打开浏览器，请在窗口里登录卖家精灵（登录态保存在 ~/.sellersprite-profile，之后免登录）。
已下载的文件会跳过，中途失败可直接重跑。失败时在输出目录生成 debug_*.png 供排查。
"""
import argparse, json, re, sys, time
from datetime import date
from decimal import Decimal
from urllib.parse import quote
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---- 页面地址（卖家精灵改版时只需改这里）----
BASE = "https://www.sellersprite.com/v3"
URL_COMPETITOR = BASE + "/competitor-lookup"
URL_EXPAND = BASE + "/traffic/extend"
URL_CONV = BASE + "/keyword-conversion-rate"
MARKET_ID = {"US": 1, "CA": 7}   # 转化率页 marketId（CA=7 已验证，US 待验证）
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

    def query(self):
        self.p.get_by_text(re.compile(r"^\s*立即查询\s*$")).locator("visible=true").first.click()
        self.p.wait_for_load_state("networkidle", timeout=60000)
        self.p.wait_for_timeout(2500)

    def quota(self):
        """悬停导出按钮，读取提示里的'当天剩余 N 次'；读不到返回 None。"""
        try:
            self.p.get_by_text(re.compile(r"^\s*导出\s*$")).locator("visible=true").first.hover()
            self.p.wait_for_timeout(800)
            m = re.search(r"当天剩余(\d+)次", self.p.locator("body").inner_text())
            return int(m.group(1)) if m else None
        except Exception: return None

    def export(self, tag):
        """点 导出，保持卖家精灵原文件名保存到结果文件夹。
        先尝试直接下载；若是后台异步导出（提示任务已添加），转到'我的导出'下载最新一条。"""
        if self.done(tag): log("已存在，跳过", self.done_map[tag]); return
        left = self.quota()
        if left is not None:
            log(f"今日导出剩余 {left} 次")
            if left <= 0: sys.exit("今日导出次数已用完，明天再跑（已下载的会自动跳过）")
        btn = self.p.get_by_role("button", name=re.compile(r"^\s*导出\s*$")).locator("visible=true")
        if btn.count() == 0:
            btn = self.p.get_by_text(re.compile(r"^\s*导出\s*$")).locator("visible=true")
        dl = None
        try:
            with self.p.expect_download(timeout=15000) as d:
                btn.first.click(); self.p.wait_for_timeout(1200); self.confirm_dialog()
            dl = d.value
        except PWTimeout:
            dl = self.download_from_my_exports(tag)
        name = dl.suggested_filename
        dl.save_as(self.out / name)
        self.done_map[tag] = name; self.mf.write_text(json.dumps(self.done_map, ensure_ascii=False, indent=1), "utf8")
        log("已保存", name)

    def confirm_dialog(self):
        for t in ("确认导出", "确认", "确定", "导出"):
            c = self.p.locator(".el-dialog:visible, .el-message-box:visible, .modal:visible, [role=dialog]:visible").get_by_text(t, exact=True)
            if c.count(): c.first.click(); return

    def download_from_my_exports(self, tag):
        """异步导出：打开'我的导出'，等最新一条完成后下载。"""
        self.shot(f"{tag}_after_export")
        link = self.p.get_by_text(re.compile("我的导出|导出记录|下载中心|下载任务")).locator("visible=true")
        if link.count() == 0: sys.exit(f"导出后未下载也找不到'我的导出'入口：{tag}（见 {self.logs} 截图，请把截图发给开发者）")
        pg = self.p
        try:                                    # 可能新开标签；没有则在原页
            with self.p.context.expect_page(timeout=5000) as np_:
                link.first.click()
            pg = np_.value
        except PWTimeout:
            pass
        for _ in range(40):                     # 最多等约 2 分钟
            pg.wait_for_timeout(3000)
            b = pg.get_by_text(re.compile(r"^\s*下载\s*$")).locator("visible=true")
            if b.count():
                with pg.expect_download(timeout=60000) as d: b.first.click()
                return d.value
            pg.reload(wait_until="domcontentloaded")
        sys.exit(f"等待导出完成超时：{tag}")

    # ---------- 竞品表（全部用网址参数，不点下拉框）----------
    def comp_url(self, kw, month):
        """month: None=近30天；'2026.07'=月表。不展示变体由页面默认（variants 未勾选）+ 下面再核对。"""
        mn = "bsr_sales_nearly" if month is None else "bsr_sales_monthly_" + month.replace(".", "")
        return (f"{URL_COMPETITOR}?market={self.site}&monthName={mn}&asins=%5B%5D&keywords={quote(kw)}"
                "&page=1&nodeIdPaths=%5B%5D&symbolFlag=false&size=60"
                "&order%5Bfield%5D=total_units&order%5Bdesc%5D=true&lowPrice=N&salesGrowthType=30d")

    def competitor(self, kw, month):
        tag = "products" if month is None else f"month:{month}"
        if self.done(tag): return
        self.p.goto(self.comp_url(kw, month), wait_until="domcontentloaded")
        self.p.wait_for_timeout(4000); self.ensure_login()
        self.settle()
        # 保证"展示所有变体"未勾选
        if self.p.locator("label.is-checked:has-text('展示所有变体')").count():
            self.p.get_by_text("展示所有变体").locator("visible=true").first.click(); self.settle()
        self.export(tag)

    def products(self, kw): self.competitor(kw, None)

    def last_months(self):
        """最近 12 个完整月份（当前月不含），如今天 2026-11 → 2025.11 … 2026.10。"""
        t = date.today(); y, m = t.year, t.month
        out = []
        for _ in range(N_MONTHS):
            m -= 1
            if m == 0: y, m = y - 1, 12
            out.append(f"{y}.{m:02d}")
        return out[::-1]

    def months(self, kw):
        for m in self.last_months(): self.competitor(kw, m)

    def settle(self):
        try: self.p.wait_for_load_state("networkidle", timeout=30000)
        except PWTimeout: pass
        self.p.wait_for_timeout(2000)

    # ---------- 头部 ASIN 关键词反查（拓展流量词）----------
    def asins(self):
        wb = openpyxl.load_workbook(self.out / self.done_map["products"], read_only=True)
        rows = wb.worksheets[0].iter_rows(values_only=True)
        head = next(rows); i = head.index("ASIN")
        out = []
        for r in rows:
            v = str(r[i]).strip() if r[i] else ""
            if re.fullmatch(r"B0[0-9A-Z]{8}|[0-9]{9}[0-9X]", v) and v not in out: out.append(v)
            if len(out) >= TOP_ASIN: break
        return out

    def keywords_table(self):
        if self.done("expand"): return
        asins = self.asins(); log("反查 ASIN：", ",".join(asins))
        (self.logs / "头部ASIN清单.txt").write_text("\n".join(asins), encoding="utf8")
        self.p.goto(f"{URL_EXPAND}?marketId={MARKET_ID[self.site]}", wait_until="domcontentloaded")
        self.p.wait_for_timeout(3000); self.ensure_login()
        self.select_site_label()
        box = self.p.locator("input:visible, textarea:visible").first
        box.click(); self.p.keyboard.insert_text(",".join(asins)); self.p.keyboard.press("Enter")
        self.p.wait_for_timeout(500)
        self.query(); self.settle()
        self.sort_traffic()
        self.export("expand")

    def select_site_label(self):
        pass  # 站点由网址 marketId 决定；若页面不认，查看 debug 截图后再补

    def sort_traffic(self):
        """结果默认已按流量占比降序；若不是，点击表头'流量占比'直到出现降序。"""
        h = self.p.get_by_text(re.compile(r"流量占比")).locator("visible=true")
        if h.count() == 0: log("未找到流量占比表头，沿用页面默认排序")

    # ---------- 转化率 ----------
    def kw_list(self):
        wb = openpyxl.load_workbook(self.out / self.done_map["expand"], read_only=True)
        rows = wb.worksheets[0].iter_rows(values_only=True)
        head = next(rows); ik, ip = head.index("关键词"), head.index("流量占比")
        cum, kws = Decimal(0), []
        for r in rows:                       # 空占比的词保留原值，不补造；按 Decimal 累计避免 95% 边界浮点误差
            v = r[ip]
            if isinstance(v, str): v = Decimal(v.strip().strip("%") or 0) / 100
            else: v = Decimal(str(v or 0))
            kws.append(str(r[ik]).strip()); cum += v
            if cum >= Decimal(str(TRAFFIC_CUM)) or len(kws) >= MAX_KW: break
        log(f"累计流量占比 {cum:.2%}，共 {len(kws)} 个关键词")
        return kws

    def conversion(self):
        if self.done("conv"): return
        kws = self.kw_list()
        (self.logs / "核心关键词清单.txt").write_text("\n".join(kws), encoding="utf8")
        url = (f"{URL_CONV}?marketId={MARKET_ID[self.site]}&reverseType=90D&bidMatchType=1"
               f"&keywordMatchType=all&matchType=1&keywordList={quote(','.join(kws), safe=',')}")
        self.p.goto(url, wait_until="domcontentloaded")
        self.p.wait_for_timeout(4000); self.ensure_login(); self.settle()
        self.export("conv")

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
    ap.add_argument("--site", default="CA", choices=["CA", "US"])
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
