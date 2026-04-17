// ==UserScript==
// @name         Yext Photo Text Generator (Description/Details/Alt Text)
// @namespace    melonlocal
// @version      0.2
// @description  Generates and fills photo Description/Details/Alt Text in Yext entity editor
// @match        https://www.yext.com/s/*/entity/edit3*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/yext-photo-text-generator.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/yext-photo-text-generator.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ----------------------------
  // Configuration
  // ----------------------------
  const CONFIG = {
    dryRun: true,                 // true = preview only; false = write to fields
    overwriteDescription: false,  // if true, replaces each photo Description with entity-level business description
    pauseMsBetweenCards: 50
  };

  // ----------------------------
  // Helpers
  // ----------------------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function normalize(s) {
    return (s || "").trim().replace(/\s+/g, " ");
  }

  // ----------------------------
  // Entity context + embedded fields
  // ----------------------------
  function getEntityContext() {
    // Best-effort from visible page content
    const firstCellName =
      normalize(document.querySelector("table td")?.textContent) ||
      normalize(document.querySelector("h1")?.textContent);

    // Try to locate address block in the Core Information table
    const tables = Array.from(document.querySelectorAll("table"));
    let cityState = "";
    for (const t of tables) {
      const txt = t.innerText || "";
      if (txt.includes("Address") && txt.match(/\b[A-Z]{2}\s+\d{5}\b/)) {
        // crude parse: look for a line like "Gainesville, GA 30501"
        const line = txt.split("\n").map(normalize).find((l) => l.match(/,\s*[A-Z]{2}\s+\d{5}\b/));
        if (line) {
          cityState = line.replace(/\s+\d{5}.*/, ""); // remove zip and after
        }
        break;
      }
    }

    // Fallback: page shows entity name plainly in header area (e.g., "John Leonard - State Farm Insurance Agent")
    const headerName = normalize(
      document.body.innerText.split("\n").find((l) => l.includes("State Farm Insurance Agent")) || ""
    );
    const finalName = headerName || firstCellName;

    // crude city/region split from "City, ST"
    let city = "";
    let region = "";
    if (cityState && cityState.includes(",")) {
      const parts = cityState.split(",");
      city = normalize(parts[0]);
      region = normalize(parts[1]);
    }

    // Website URL best-effort (assumes visible in page text somewhere)
    const bodyText = document.body.innerText || "";
    const websiteMatch = bodyText.match(/https?:\/\/[^\s]+/);
    const websiteUrl = websiteMatch ? normalize(websiteMatch[0]) : "";

    return {
      rawName: finalName || "",
      name: finalName || "This business",
      cityState: cityState || "",
      city,
      region,
      websiteUrl
    };
  }

  function buildEmbeddedFieldMap(ctx) {
    // Keys are literal substrings on the page; values are Yext embedded tokens.
    // Extend this as needed with more entity fields.
    const map = new Map();

    if (ctx.rawName) {
      map.set(ctx.rawName, "[[name]]");
    }
    if (ctx.name && ctx.name !== ctx.rawName) {
      map.set(ctx.name, "[[name]]");
    }
    if (ctx.city) {
      map.set(ctx.city, "[[address.city]]");
    }
    if (ctx.region) {
      map.set(ctx.region, "[[address.region]]");
    }
    if (ctx.city && ctx.region) {
      map.set(`${ctx.city}, ${ctx.region}`, "[[address.city]], [[address.region]]");
    }
    if (ctx.websiteUrl) {
      map.set(ctx.websiteUrl, "[[website.url]]");
    }

    return map;
  }

  function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function applyLiteralToTokenMap(input, map) {
    if (!input) return input;
    let out = input;
    // Replace longer literals first to avoid partial overlaps.
    const entries = Array.from(map.entries()).sort((a, b) => b[0].length - a[0].length);
    for (const [literal, token] of entries) {
      if (!literal) continue;
      const re = new RegExp(escapeRegExp(literal), "g");
      out = out.replace(re, token);
    }
    return out;
  }

  function isSystemyDescription(desc, ctx) {
    const s = normalize(desc).toLowerCase();
    if (!s) return true;
    if (s === "[[name]]" || s === "[[website.url]]") return true;
    if (ctx.rawName && s === `${ctx.rawName.toLowerCase()} photo`) return true;
    if (ctx.name && s === `${ctx.name.toLowerCase()} photo`) return true;
    if (/photo$/i.test(s) && s.length <= 80 && s.includes("state farm")) return true;
    return false;
  }

  function buildDescription(ctx, existingDesc, tokenMap) {
    const desc = normalize(existingDesc);

    // If already tokenized and not "systemy", just normalize spacing and return.
    if (/\[\[.+?\]\]/.test(desc) && !isSystemyDescription(desc, ctx)) {
      return desc;
    }

    // Systemy or empty: synthesize "<rawName> photo" then tokenize to [[name]] photo
    if (isSystemyDescription(desc, ctx)) {
      const base = ctx.rawName || ctx.name || "This business";
      let generated = `${base} photo`;
      generated = applyLiteralToTokenMap(generated, tokenMap);
      return generated;
    }

    // Human-written caption: keep semantics, just token-replace literals.
    const tokenized = applyLiteralToTokenMap(desc, tokenMap);
    return tokenized;
  }

  function buildBusinessDescription(ctx) {
    // Evergreen, publisher-safe, not promotional
    const where = ctx.cityState ? ` in ${ctx.cityState}` : "";
    return `${ctx.name}${where} helps customers with auto, home, renters, life, and business insurance, focusing on clear guidance and long-term relationships.`;
  }

  function buildDetails(ctx, seed) {
    const s = (seed || "").toLowerCase();

    if (!seed) {
      return `${ctx.name} sharing an update from the agency and local community.`;
    }

    if (s.includes("birthday")) {
      return `${ctx.name} attending a birthday lunch while connecting with customers and community members.`;
    }

    if (s.includes("lake") || s.includes("boat")) {
      return `${ctx.name} enjoying time on the lake, representing coverage for boats and recreational assets.`;
    }

    if (s.includes("jeep")) {
      return `${ctx.name} highlighting auto insurance coverage options for Jeep owners.`;
    }

    if (s.includes("classic") || s.includes("bugs") || s.includes("beetle")) {
      return `A classic vehicle representing classic car insurance options available through the ${ctx.name} agency.`;
    }

    // Default: contextualize the seed as a caption
    return `${ctx.name} sharing a moment from the agency: ${seed}.`;
  }

  function buildAltText(ctx, seed) {
    // Literal and short; avoid marketing language
    if (!seed) return `${ctx.name} in the office.`;

    const s = (seed || "").toLowerCase();
    if (s.includes("birthday")) return `${ctx.name} at a birthday lunch.`;
    if (s.includes("lake") || s.includes("boat")) return `${ctx.name} on a boat on a lake.`;
    if (s.includes("jeep")) return `Jeep vehicle, representing auto insurance.`;
    if (s.includes("classic") || s.includes("bugs") || s.includes("beetle")) return `Classic Volkswagen Beetle.`;

    // Default: reuse seed but keep it literal
    return `${ctx.name}: ${seed}.`.replace(/!+/g, ".").replace(/\s+/g, " ").trim();
  }

  function findPhotoCards() {
    // A "card" here is any ancestor container that contains all three fields.
    const descAreas = Array.from(document.querySelectorAll('textarea[aria-label="description"]'));
    const cards = [];

    for (const desc of descAreas) {
      const card = desc.closest("td, div, section, article") || desc.parentElement;
      if (!card) continue;

      const details = card.querySelector('textarea[aria-label="details"]');
      const alt = card.querySelector('textarea[aria-label="alternateText"]');
      if (details && alt) cards.push({ card, desc, details, alt });
    }

    return cards;
  }

  function setTextareaValue(textarea, value) {
    textarea.focus();
    textarea.value = value;

    // Trigger React-controlled input listeners
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
    textarea.blur();
  }

  // ----------------------------
  // UI Button
  // ----------------------------
  function addButton() {
    const btn = document.createElement("button");
    btn.textContent = CONFIG.dryRun ? "Preview Photo Text" : "Generate Photo Text";
    btn.style.position = "fixed";
    btn.style.bottom = "16px";
    btn.style.right = "16px";
    btn.style.zIndex = "99999";
    btn.style.padding = "10px 12px";
    btn.style.borderRadius = "10px";
    btn.style.border = "1px solid #ccc";
    btn.style.background = "#fff";
    btn.style.cursor = "pointer";
    btn.style.boxShadow = "0 4px 18px rgba(0,0,0,0.12)";

    btn.addEventListener("click", async () => {
      const ctx = getEntityContext();
      const tokenMap = buildEmbeddedFieldMap(ctx);
      const businessDesc = buildBusinessDescription(ctx);

      const cards = findPhotoCards();
      if (!cards.length) {
        alert("No photo cards found. Scroll to the Photo Gallery section and try again.");
        return;
      }

      const preview = [];
      for (const [i, c] of cards.entries()) {
        const seed = normalize(c.desc.value);

        const generatedDesc = CONFIG.overwriteDescription
          ? businessDesc
          : buildDescription(ctx, seed, tokenMap);

        const newDetails = buildDetails(ctx, seed);
        const newAlt = buildAltText(ctx, seed);

        preview.push({
          i: i + 1,
          seed,
          description: generatedDesc,
          details: newDetails,
          altText: newAlt
        });

        if (!CONFIG.dryRun) {
          if (CONFIG.overwriteDescription) {
            setTextareaValue(c.desc, generatedDesc);
          } else if (generatedDesc !== seed) {
            setTextareaValue(c.desc, generatedDesc);
          }
          setTextareaValue(c.details, newDetails);
          setTextareaValue(c.alt, newAlt);
          await sleep(CONFIG.pauseMsBetweenCards);
        }
      }

      if (CONFIG.dryRun) {
        console.table(preview);
        alert("Preview complete. Open the console to view the table (Right click → Inspect → Console).");
      } else {
        alert(`Done. Updated ${cards.length} photo(s). Review changes, then save in Yext.`);
      }
    });

    document.body.appendChild(btn);
  }

  addButton();
})();