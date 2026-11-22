import re
from playwright.sync_api import sync_playwright

BASE = "https://www.boclife.com.hk/tc/support/financial-information/"
PAGES = [("分紅實現率", BASE + "dividend-fulfillment-ratios.html"),
         ("總現金價值比率", BASE + "total-cash-value-ratio.html")]

JS_TABLES = """() => Array.from(document.querySelectorAll('table')).map(t => {
  const rows = Array.from(t.rows).map(tr => Array.from(tr.cells).map(
    c => (c.innerText || c.textContent || '').replace(/\\s+/g, ' ').trim()));
  const r = t.getBoundingClientRect();
  return {rows: rows, visible: r.width > 0 && r.height > 0};
})"""
JS_SELECTS = """() => Array.from(document.querySelectorAll('select')).map(s => ({
  name: s.name || s.id || s.className || '',
  options: Array.from(s.options).map(o => (o.textContent || '').trim())
}))"""
JS_CONTROLS = """() => Array.from(document.querySelectorAll(
  '[class*="tab"],[class*="select"],[class*="dropdown"],[class*="accordion"],[class*="collapse"],[role="tab"],[role="button"]'))
  .map(e => ({cls: String(e.className || '').slice(0, 60),
              txt: (e.innerText || e.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 80)}))
  .filter(x => x.txt)"""
JS_LINKS = """() => Array.from(document.querySelectorAll('a[href]')).map(a => ({
  href: a.href, txt: (a.innerText || a.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 80)}))"""
JS_FRAMES = "() => Array.from(document.querySelectorAll('iframe')).map(f => f.src || '')"
JS_SCRIPTS = """() => Array.from(document.querySelectorAll('script')).map(
  s => s.src ? '' : (s.textContent || ''))"""
JS_TEXT = "() => document.body ? document.body.innerText : ''"

PCT = re.compile(r"\d\s*[%％]")
NA = re.compile(r"不適用|不适用|N/?A")
KEY = re.compile(r"實現率|比率|報告年度|保單年度|貨幣|註|不適用")
FILE = re.compile(r"\.(pdf|xlsx?|csv|json)(\?|$)", re.I)


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


def probe(ctx, label, url):
    print("\n==================== %s ====================" % label)
    print("網址：%s" % url)
    page = ctx.new_page()
    seen = []

    def on_response(resp):
        try:
            rtype = resp.request.resource_type
            ctype = resp.headers.get("content-type", "")
            if rtype in ("xhr", "fetch") or "json" in ctype:
                try:
                    body = resp.text()
                except Exception:
                    body = ""
                seen.append((resp.request.method, resp.status, resp.url, ctype, body,
                             resp.request.post_data or ""))
        except Exception:
            pass

    page.on("response", on_response)
    status = "?"
    try:
        r = page.goto(url, wait_until="domcontentloaded", timeout=60000)
        status = r.status if r else "?"
    except Exception as e:
        print("載入出錯：%s" % one_line(str(e))[:200])
    try:
        page.wait_for_load_state("networkidle", timeout=30000)
    except Exception:
        pass
    page.wait_for_timeout(4000)

    try:
        text = page.evaluate(JS_TEXT)
        print("\n=== 1. 載入 ===")
        print("   狀態 %s｜標題：%s｜文字 %d 字｜百分比 %d 個｜不適用 %d 個" % (
            status, one_line(page.title())[:80], len(text),
            len(PCT.findall(text)), len(NA.findall(text))))
        years = sorted(set(re.findall(r"20\d{2}", text)))
        print("   出現的年份：%s" % " ".join(years[:30]))

        print("\n=== 2. 表格 ===")
        tables = page.evaluate(JS_TABLES)
        print("   共 %d 個" % len(tables))
        for i, t in enumerate(tables[:40], 1):
            rows = t["rows"]
            cells = [c for row in rows for c in row]
            print("   表格 %d｜%s｜%d 行｜最多 %d 欄｜百分比 %d｜不適用 %d" % (
                i, "可見" if t["visible"] else "隱藏", len(rows),
                max([len(x) for x in rows] or [0]),
                sum(1 for c in cells if PCT.search(c)), sum(1 for c in cells if NA.search(c))))
            for row in rows[:6]:
                print("      %s" % "｜".join(row)[:300])
            if len(rows) > 6:
                print("      …最後一行：%s" % "｜".join(rows[-1])[:300])

        print("\n=== 3. 下拉選單 ===")
        for s in page.evaluate(JS_SELECTS)[:10]:
            print("   %s｜%d 項：%s" % (s["name"][:40], len(s["options"]),
                                       " ｜ ".join(s["options"])[:600]))

        print("\n=== 4. 分頁／摺疊等控制項 ===")
        shown = []
        for c in page.evaluate(JS_CONTROLS):
            line = "%s｜%s" % (c["cls"], c["txt"])
            if line not in shown:
                shown.append(line)
        print("   共 %d 個（最多顯示 40 個）" % len(shown))
        for x in shown[:40]:
            print("   %s" % x)

        print("\n=== 5. 檔案及相關連結 ===")
        links = []
        for a in page.evaluate(JS_LINKS):
            if FILE.search(a["href"]) or re.search(r"實現率|比率", a["txt"]):
                line = "%s｜%s" % (a["txt"], a["href"])
                if line not in links:
                    links.append(line)
        print("   共 %d 個（最多顯示 60 個）" % len(links))
        for x in links[:60]:
            print("   %s" % x[:300])

        print("\n=== 6. iframe ===")
        for f in page.evaluate(JS_FRAMES)[:10]:
            print("   %s" % f[:300])

        print("\n=== 7. 網頁發出的資料請求 ===")
        print("   共 %d 個（最多顯示 30 個）" % len(seen))
        for method, st, u, ctype, body, post in seen[:30]:
            print("   %s %s %s｜%s｜%d 字｜百分比 %d" % (
                method, st, u[:200], ctype[:40], len(body), len(PCT.findall(body))))
            if post:
                print("      請求內容：%s" % one_line(post)[:300])
            if PCT.search(body) or NA.search(body):
                print("      回應開頭：%s" % one_line(body)[:1200])

        print("\n=== 8. 內嵌 script 內的資料 ===")
        for i, s in enumerate(page.evaluate(JS_SCRIPTS), 1):
            n = len(PCT.findall(s))
            if n >= 3:
                pos = PCT.search(s).start()
                print("   script %d｜%d 字｜百分比 %d｜片段：%s" % (
                    i, len(s), n, one_line(s[max(0, pos - 600):pos + 600])))

        print("\n=== 9. 頁面相關文字（最多 60 行） ===")
        lines = [one_line(x) for x in text.split("\n") if KEY.search(x) and len(x.strip()) > 3]
        for x in lines[:60]:
            print("   %s" % x[:300])
    except Exception as e:
        print("探查出錯：%s" % one_line(str(e))[:300])
    page.close()


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        try:
            ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                      viewport={"width": 1366, "height": 900})
            for label, url in PAGES:
                probe(ctx, label, url)
        finally:
            browser.close()
    print("\n=== 完 ===")


if __name__ == "__main__":
    main()
