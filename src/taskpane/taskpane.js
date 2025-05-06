/* ------------------------------------------------------------------
 *  OneNote Counter – recursive Office JS implementation
 *  No Microsoft Graph / no SSO required
 * -----------------------------------------------------------------*/
/* global Office, OneNote */

// ‑‑‑ UI bootstrap
Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

/* ------------------------------------------------------------------
 *  PARENT → CHILD map for validation
 * -----------------------------------------------------------------*/
const CHILDREN = {
  Application: ["Notebook", "Section Group", "Section", "Page", "SubPage"],
  "Notebook": ["Section Group", "Section", "Page", "SubPage"],
  "Section Group": ["Section Group", "Section", "Page", "SubPage"],
  Section: ["Page", "SubPage"],
  Page: ["Word count"],
  SubPage: ["Word count"],
};

/* ------------------------------------------------------------------
 *  Entry‑point wired to the “Run” button
 * -----------------------------------------------------------------*/
export async function run() {
  const scopeSel = document.getElementById("scope-dropdown").value;   // first drop‑down
  const itemSel  = document.getElementById("count-dropdown").value;   // second drop‑down
  const resultBox = document.getElementById("result");

  /* 1. validate parent‑child relationship */
  if (!CHILDREN[scopeSel] || !CHILDREN[scopeSel].includes(itemSel)) {
    resultBox.innerText = `❌  “${itemSel}” isn’t inside “${scopeSel}”.`;
    return;
  }

  try {
    await OneNote.run(async (context) => {
      /* 2. get the starting scope object */
      const scopeObj = await getScopeObject(context, scopeSel);

      /* 3. handle Word‑count separately */
      if (itemSel === "Word count") {
        const words = await countWordsOnPage(context, scopeObj);
        resultBox.innerText = `📝  Word count on this page: ${words.toLocaleString()}`;
        return;
      }

      /* 4. recursive enumeration for Notebook / SectionGroup / Section / Page / SubPage */
      const total = await countEntities(context, scopeObj, itemSel);
      resultBox.innerText = `✅  ${itemSel}${itemSel === "Page" ? "s" : ""} in ${scopeSel}: ${total.toLocaleString()}`;
    });
  } catch (err) {
    console.error(err);
    resultBox.innerText = `❌  Error: ${err.message}`;
  }
}

/* ------------------------------------------------------------------
 *  Returns the ClientObject that matches the first drop‑down choice
 * -----------------------------------------------------------------*/
async function getScopeObject(ctx, scopeName) {
  switch (scopeName) {
    case "Application":
      return ctx.application;

    case "Notebook": {
      const nb = ctx.application.getActiveNotebookOrNull();
      nb.load("id");     // ensure it’s not null
      await ctx.sync();
      if (!nb.id) throw new Error("No active notebook.");
      return nb;
    }

    case "Section Group": {
      const sg = ctx.application.getActiveSectionGroupOrNull();
      sg.load("id");
      await ctx.sync();
      if (!sg.id) throw new Error("No active section group.");
      return sg;
    }

    case "Section": {
      const sec = ctx.application.getActiveSectionOrNull();
      sec.load("id");
      await ctx.sync();
      if (!sec.id) throw new Error("No active section.");
      return sec;
    }

    case "Page":
    case "SubPage": {
      const pg = ctx.application.getActivePageOrNull();
      pg.load("id");
      await ctx.sync();
      if (!pg.id) throw new Error("No active page.");
      return pg;
    }

    default:
      throw new Error(`Unknown scope: ${scopeName}`);
  }
}

/* ------------------------------------------------------------------
 *  Generic recursive counter
 * -----------------------------------------------------------------*/
async function countEntities(ctx, scope, target) {
  let total = 0;

  /* ---- load direct collections we need on this scope ---- */
  if (target === "Notebook" && scope.notebooks)    scope.notebooks.load("items");
  if (scope.sectionGroups)                         scope.sectionGroups.load("items");
  if (scope.sections)                              scope.sections.load("items");

  if ((target === "Page" || target === "SubPage") && scope.pages)
    scope.pages.load("items,level");

  await ctx.sync();

  /* ---- 1. direct matches at this level ------------------ */
  if (target === "Notebook" && scope.notebooks)
    total += scope.notebooks.items.length;

  if (target === "Section Group" && scope.sectionGroups)
    total += scope.sectionGroups.items.length;

  if (target === "Section" && scope.sections)
    total += scope.sections.items.length;

  if ((target === "Page" || target === "SubPage") && scope.pages) {
    for (const page of scope.pages.items) {
      if (target === "Page")                       total += 1;
      else if (page.level > 1)                     total += 1;   // SubPage = page.level > 1
    }
  }

  /* ---- 2. recurse into children ------------------------- */
  if (scope.sectionGroups) {
    for (const sg of scope.sectionGroups.items) {
      total += await countEntities(ctx, sg, target);
    }
  }

  if (scope.sections) {
    for (const sec of scope.sections.items) {
      total += await countEntities(ctx, sec, target);
    }
  }

  // notebooks cannot nest; pages cannot have children – recursion ends.
  return total;
}

/* ------------------------------------------------------------------
 *  Word‑count helper – only valid when scopeObj is Page / SubPage
 * -----------------------------------------------------------------*/
async function countWordsOnPage(ctx, pageObj) {
  const htmlResult = pageObj.getHtml();
  await ctx.sync();
  const html = htmlResult.value || "";

  // strip tags, collapse whitespace, count words
  const textOnly = html.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
  return textOnly ? textOnly.split(" ").length : 0;
}
