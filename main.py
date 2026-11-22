import collections
import re
from playwright.sync_api import sync_playwright

BASE = "https://www.boclife.com.hk/tc/support/financial-information/"
PAGES = [("分紅實現率", BASE + "dividend-fulfillment-ratios.html"),
         ("總現金價值比率", BASE + "total-cash-value-ratio.html")]

JS_TABLES = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  const tag = e => e.tagName.toLowerCase() + (e.id ? '#' + e.id : '') +
    (typeof e.className === 'string' && e.className.trim()
      ? '.' + e.className.trim().split(/\\s+/).slice(0, 3).join('.') : '');
  return Array.from(document.querySelectorAll('table')).map(t => {
    const rows = Array.from(t.rows).map(tr => Array.from(tr.cells).map(c => clean(c.textContent)));
    const path = [], ctx = [];
    let node = t;
    for (let d = 1; d <= 9 && node.parentElement && node.parentElement !== document.body; d++) {
      let sib = node.previousElementSibling;
      for (let k = 0; sib && k < 3; k++) {
        const s = clean(sib.textContent);
        if (s && s.length <= 160 && sib.tagName !== 'TABLE' && !sib.querySelector('table'))
          ctx.push('L' + d + ' ' + tag(sib) + '：' + s);
        sib = sib.previousElementSibling;
      }
      node = node.parentElement;
      path.push(tag(node));
      if (node.id) {
        const id = node.id;
        const sel = '[href="#' + id + '"],[data-bs-target="#' + id + '"],[data-target="#' + id +
                    '"],[aria-controls="' + id + '"]';
        try {
          document.querySelectorAll(sel).forEach(x => {
            const s = clean(x.textContent);
            if (s && s.length <= 160) ctx.push('L' + d + ' 控制 #' + id + '：' + s);
          });
        } catch (e) {}
      }
    }
    return {rows: rows, path: path, ctx: ctx};
  });
}"""
JS_HTML = """([i, depth]) => {
  let n = document.querySelectorAll('table')[i];
  if (!n) return '';
  for (let d = 0; d < depth && n.parentElement && n.parentElement !== document.body; d++)
    n = n.parentElement;
  const c = n.cloneNode(true);
  c.querySelectorAll('table').forEach(x => x.replaceWith(document.createTextNode('[TABLE]')));
  return c.outerHTML.replace(/\\s+/g, ' ');
}"""
JS_NOTES = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  const out = [];
  document.querySelectorAll('p, li, span, div, small').forEach(e => {
    if (e.closest('table') || e.children.length > 3) return;
    const s = clean(e.textContent);
    if (s.length >= 8 && s.length <= 500 &&
        /N[/]A|不適用|尚未推出|已停售|註|報告年度|貨幣|港元|美元|人民幣/.test(s) && !out.includes(s))
      out.push(s);
  });
  return out;
}"""

PCT = re.compile(r"^\d+(?:\.\d+)?\s*[%％]$")


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


def is_desktop(t):
    return max([len(r) for r in t["rows"]] or [0]) >= 11


def values(t):
    rows = t["rows"]
    if is_desktop(t):
        return rows[-1][1:] if rows else []
    return [r[-1] for r in rows[1:]]


def probe(ctx, label, url):
    print("\n==================== %s ====================" % label)
    page = ctx.new_page()
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        print("載入出錯：%s" % one_line(str(e))[:200])
    try:
        page.wait_for_selector("table", timeout=45000, state="attached")
    except Exception:
        print("等不到表格")
    page.wait_for_timeout(4000)
    try:
        tables = page.evaluate(JS_TABLES)
        desk = [i for i, t in enumerate(tables) if is_desktop(t)]
        print("\n=== 1. 表格總數 %d｜橫排（桌面版）%d｜其他 %d ===" % (
            len(tables), len(desk), len(tables) - len(desk)))

        print("\n=== 2. 桌面版對手機版 ===")
        same = diff = 0
        for i in desk:
            if i + 1 < len(tables) and not is_desktop(tables[i + 1]):
                a, b = values(tables[i]), values(tables[i + 1])
                if a == b:
                    same += 1
                else:
                    diff += 1
                    if diff <= 8:
                        print("   表格 %d 不同｜桌面 %s｜手機 %s" % (i + 1, " ".join(a), " ".join(b)))
        print("   相同 %d 組｜不同 %d 組" % (same, diff))

        print("\n=== 3. 非百分比的格 ===")
        cnt = collections.Counter()
        for i in desk:
            for v in values(tables[i]):
                if not PCT.match(v):
                    cnt[v[:60]] += 1
        for k, n in cnt.most_common(20):
            print("   %s × %d" % (k, n))
        cats = collections.Counter(tables[i]["rows"][-1][0] for i in desk if tables[i]["rows"])
        print("   第一欄（產品類別）：%s" % "｜".join("%s × %d" % x for x in cats.items()))

        print("\n=== 4. 每個桌面版表格的上文 ===")
        last_path = None
        for i in desk:
            t = tables[i]
            if t["path"] != last_path:
                last_path = t["path"]
                print("   [路徑] %s" % " < ".join(t["path"])[:500])
            print("   表格 %d｜%s" % (i + 1, " ".join(values(t))[:200]))
            for line in t["ctx"][:8]:
                print("      %s" % line[:260])

        print("\n=== 5. 頁面結構樣本（表格以 [TABLE] 代替） ===")
        for i in ([desk[0], desk[-1]] if desk else []):
            for depth in (3, 6):
                html = page.evaluate(JS_HTML, [i, depth])
                print("   表格 %d 往上 %d 層（%d 字）：%s" % (i + 1, depth, len(html), html[:2500]))

        print("\n=== 6. 表格以外的相關文字（最多 40 行） ===")
        for s in page.evaluate(JS_NOTES)[:40]:
            print("   %s" % s[:400])
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
