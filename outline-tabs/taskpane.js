/* Outline Tabs: Google Docs-style heading outline for Word */
let headings = [];          // [{ idx, level, text }]
let currentIdx = -1;        // paragraph index of the section the cursor is in
let collapsed = new Set();  // keys "idx" of collapsed headings
let refreshTimer = null;
let busy = false;

const $ = (id) => document.getElementById(id);

function loadCollapsed() {
  try { collapsed = new Set(JSON.parse(localStorage.getItem("ot-collapsed") || "[]")); } catch (e) {}
}
function saveCollapsed() {
  try { localStorage.setItem("ot-collapsed", JSON.stringify([...collapsed])); } catch (e) {}
}

// Map a paragraph to a heading level 0-5, or -1 if not a heading.
function levelOf(p) {
  const sb = p.styleBuiltIn;
  if (sb === Word.Style.title) return 0;
  const map = {
    [Word.Style.heading1]: 0, [Word.Style.heading2]: 1, [Word.Style.heading3]: 2,
    [Word.Style.heading4]: 3, [Word.Style.heading5]: 4, [Word.Style.heading6]: 5,
  };
  if (sb in map) return map[sb];
  const m = /^heading\s*([1-6])$/i.exec(p.style || "");
  return m ? Number(m[1]) - 1 : -1;
}

async function refresh() {
  if (busy) return;
  busy = true;
  try {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text,items/style,items/styleBuiltIn");
      const sel = ctx.document.getSelection();
      await ctx.sync();

      const found = [];
      paras.items.forEach((p, i) => {
        const lvl = levelOf(p);
        const text = (p.text || "").trim();
        if (lvl >= 0 && text) found.push({ idx: i, level: lvl, text, para: p });
      });

      // Which heading is the cursor under? Last heading starting at or before selection.
      const cmp = found.map((h) => h.para.getRange("Start").compareLocationWith(sel));
      await ctx.sync();
      let cur = -1;
      cmp.forEach((c, i) => {
        // heading start is Before / AdjacentBefore / Inside / Equal relative to the selection
        if (["Before", "AdjacentBefore", "Equal", "Contains", "Inside", "OverlapsStart"].includes(c.value)) cur = found[i].idx;
      });

      headings = found.map(({ idx, level, text }) => ({ idx, level, text }));
      currentIdx = cur;
    });
    render();
  } catch (e) {
    console.error(e);
  } finally {
    busy = false;
  }
}

function scheduleRefresh() {
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(refresh, 500);
}

function hasChildren(i) {
  return i + 1 < headings.length && headings[i + 1].level > headings[i].level;
}

function render() {
  const nav = $("outline");
  const q = $("filter").value.trim().toLowerCase();
  nav.textContent = "";
  $("empty").hidden = headings.length > 0;

  // base level so the top-most heading is flush left
  const base = headings.length ? Math.min(...headings.map((h) => h.level)) : 0;
  let hideBelow = null; // level of a collapsed ancestor

  headings.forEach((h, i) => {
    if (hideBelow !== null) {
      if (h.level > hideBelow) return;
      hideBelow = null;
    }
    if (q && !h.text.toLowerCase().includes(q)) return;

    const row = document.createElement("div");
    row.className = `row lvl-${h.level}` + (h.idx === currentIdx ? " current" : "");
    row.style.paddingLeft = (h.level - base) * 14 + 4 + "px";

    const tw = document.createElement("button");
    tw.className = "twisty" + (hasChildren(i) && !q ? "" : " empty");
    const isCollapsed = collapsed.has(String(h.idx));
    tw.textContent = isCollapsed ? "\u25B6" : "\u25BC";
    tw.setAttribute("aria-label", isCollapsed ? "Expand" : "Collapse");
    tw.onclick = (ev) => {
      ev.stopPropagation();
      const k = String(h.idx);
      collapsed.has(k) ? collapsed.delete(k) : collapsed.add(k);
      saveCollapsed();
      render();
    };

    const label = document.createElement("div");
    label.className = "label";
    label.textContent = h.text;
    label.title = h.text;
    label.onclick = () => jumpTo(h);

    row.append(tw, label);
    nav.append(row);

    if (isCollapsed && hasChildren(i) && !q) hideBelow = h.level;
  });
}

async function jumpTo(h) {
  try {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text");
      await ctx.sync();
      let target = paras.items[h.idx];
      // Document may have shifted since last refresh; fall back to a text match.
      if (!target || (target.text || "").trim() !== h.text) {
        target = paras.items.find((p) => (p.text || "").trim() === h.text);
      }
      if (target) target.select("Start");
      await ctx.sync();
    });
    currentIdx = h.idx;
    render();
  } catch (e) {
    console.error(e);
  }
}

function setAll(collapse) {
  collapsed = new Set();
  if (collapse) headings.forEach((h, i) => { if (hasChildren(i)) collapsed.add(String(h.idx)); });
  saveCollapsed();
  render();
}

Office.onReady(() => {
  loadCollapsed();
  $("refresh").onclick = refresh;
  $("filter").oninput = render;
  $("expandAll").onclick = () => setAll(false);
  $("collapseAll").onclick = () => setAll(true);

  // Follow the cursor and edits.
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefresh);
  try {
    Word.run(async (ctx) => {
      ctx.document.body.onParagraphChanged?.add(async () => scheduleRefresh());
      await ctx.sync();
    });
  } catch (e) { /* older Word: selection events + manual refresh still work */ }
  setInterval(scheduleRefresh, 5000); // safety net

  refresh();
});
