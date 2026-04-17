// ==UserScript==
// @name         Yext Media Autofill
// @namespace    melonlocal
// @version      24.0
// @description  Generates unique SEO-focused Description, Details, and Alt Text for each Yext media item via Claude or Gemini vision APIs.
// @match        https://www.yext.com/s/*/entity/edit3*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      127.0.0.1
// @connect      localhost
// @connect      generativelanguage.googleapis.com
// @connect      api.anthropic.com
// @connect      *
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/yext-media-autofill.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/yext-media-autofill.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ===========================
  // CONFIG
  // ===========================
  const CONFIG = {
    localCaptionUrl: "http://127.0.0.1:5005/caption",

    // API keys must be set via the browser console:
    //   GM_setValue("geminiApiKey", "<your-key>")
    //   GM_setValue("anthropicApiKey", "<your-key>")
    // No embedded fallbacks — keys committed in source are leaks.
    getGeminiApiKey: () => GM_getValue("geminiApiKey", ""),
    getAnthropicApiKey: () => GM_getValue("anthropicApiKey", ""),

    // Provider selection:
    //   "claude"  — Claude for all fields
    //   "gemini"  — Gemini for all fields
    //   "hybrid"  — Claude for description + details (nuanced), Gemini for alt (short)
    provider: "claude",
    claudeModel: "claude-haiku-4-5",
    geminiModel: "gemini-2.5-flash-lite",

    setClickthroughUrl: true,
    clickthroughValue: "[[website.url]]",

    minChars: 5,
    maxChars: 249,

    pauseMsBetweenItems: 120,
    apiRetryAttempts: 3,
    apiRetryBaseDelay: 1000,

    // Number of items processed in parallel.
    itemConcurrency: 3,

    defaultMode: "suggest",

    maxDim: 1024,
    jpegQuality: 0.85,
    maxDataUrlChars: 4_000_000,

    altMode: "vision",

    selectors: {
      description: 'textarea[aria-label="description"]',
      details: 'textarea[aria-label="details"]',
      alternateText: 'textarea[aria-label="alternateText"]',
      clickthroughUrl: 'input[aria-label="clickthroughUrl"]',
      fallbackDescription: 'textarea[name*="description"]',
      fallbackDetails: 'textarea[name*="details"]',
      fallbackAlt: 'textarea[name*="alternateText"]'
    }
  };

  const PROMO_SIGNALS = [
    "call today", "call us", "visit us", "stop by", "get started",
    "catch it live", "learn more", "schedule", "book now", "click",
    "we're happy", "we are happy", "ready with", "here to help",
    "answer questions", "happy new year", "as you plan", "guidance",
    "coverage options", "driving our way", "got competitive", "let's go",
    "so excited", "so proud", "check it out", "don't miss",
    "mark your calendar", "join us", "come see", "we had a blast",
    "had a great time", "what a great", "what an amazing", "shoutout",
    "big news", "exciting news", "proud to announce", "thrilled to",
    "honored to", "swipe up", "link in bio", "dm us", "tag a friend",
    "double tap", "like and share", "follow us", "stay tuned",
    "last month", "last week", "last year",
    "i was able to", "i had the pleasure", "i got to",
    "thank you for having", "for having us", "doing amazing things",
    "making a positive impact", "making a difference", "these agents",
    "my new friends", "hang with", "let our team", "our team",
    "help you find", "find the right", "insurance coverage today",
    "insurance needs", "ready to assist", "here to assist",
    "ready to help", "protect what matters", "peace of mind",
    "get a quote", "get covered", "coverage today", "right coverage",
    "insurance solutions", "serving your needs", "we can help", "i can help"
  ];

  const escapeRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  // Single compiled alternation — one pass over the input instead of N substring checks.
  const PROMO_REGEX = new RegExp(PROMO_SIGNALS.map(escapeRegex).join("|"), "i");

  const EMBEDDED_TOKENS = [
    "[[additionalHoursText]]", "[[address.extraDescription]]",
    "[[address.line1]]", "[[address.line2]]", "[[alternatePhone]]",
    "[[address.city]]", "[[countryCode]]", "[[covid19InformationUrl]]",
    "[[description]]", "[[emails]]", "[[entityId]]", "[[fax]]",
    "[[featuredMessage.description]]", "[[featuredMessage.url]]",
    "[[geomodifier]]", "[[googlePlaceId]]", "[[keywords]]",
    "[[landingPageUrl]]", "[[localeCode]]", "[[menu.displayUrl]]",
    "[[menu.url]]", "[[mobilePhone]]", "[[name]]", "[[neighborhood]]",
    "[[order.displayUrl]]", "[[order.url]]", "[[mainPhone]]",
    "[[phoneticName]]", "[[address.postalCode]]", "[[reservation.displayUrl]]",
    "[[reservation.url]]", "[[services]]", "[[slug]]",
    "[[address.region]]", "[[address.sublocality]]", "[[tollFreePhone]]",
    "[[ttyPhone]]", "[[website]]", "[[website.displayUrl]]",
    "[[website.url]]", "[[what3WordsAddress]]", "[[yextId]]",
    "[[videos.url]]"
  ];

  const FORBIDDEN_OUTPUT_TOKENS = EMBEDDED_TOKENS.filter((t) => {
    const x = (t || "").toLowerCase();
    return x.includes(".url") || x.includes("displayurl") || x === "[[website]]";
  });

  // Static system prompt — shared across all Claude calls in a batch, marked with
  // cache_control so Anthropic's prompt cache can serve it at ~10% cost on repeat
  // calls. Caching activates when the rendered prefix hits the model's minimum
  // (4096 tokens for Haiku 4.5, 2048 for Sonnet 4.6). The marker is a no-op if the
  // prompt is too short — no error, just no cache write.
  const SYSTEM_PROMPT = [
    "You generate short, factual, SEO-focused descriptive text for business photo metadata fields in Yext.",
    "",
    "OUTPUT RULES — NEVER VIOLATE:",
    "- Length: 5 to 249 characters. Never shorter, never longer.",
    "- No newlines, paragraphs, or line breaks of any kind.",
    "- Plain ASCII hyphens only. No em-dashes or en-dashes.",
    "- No double quotes, smart quotes, or backticks. Straight apostrophes (') are fine.",
    "- No slashes (forward or backward), hashtags (#), or emoji.",
    "- No phone numbers, email addresses, street addresses, or ZIP codes.",
    "- No promotional language, calls to action, or social-media phrasing. Banned phrases include but are not limited to: call today, visit us, stop by, click, learn more, book now, get started, check it out, don't miss, follow us, tag a friend, link in bio, stay tuned, swipe up, dm us, double tap, proud to announce, thrilled to, so excited, had a great time, what a great, what an amazing, shoutout, big news, exciting news, honored to, mark your calendar, join us, come see, we had a blast, doing amazing things, making a difference, ready to help, protect what matters, peace of mind, get a quote, get covered, right coverage, insurance solutions, serving your needs.",
    "- No first-person social captions (I had the pleasure, we had a blast, I was able to).",
    "",
    "FIELD TYPES:",
    "- description: 5 to 249 chars. Descriptive and factual. Unique to what's in this specific photo.",
    "- details: 5 to 249 chars. Different from the description — add context, mood, or a complementary detail. Do not simply restate the description.",
    "- alt: 5 to 50 chars preferred (hard max 249). Literal and short. Focus on what is visually in the frame, not interpretation.",
    "- business_description: 80 to 200 chars. Evergreen description of the business overall, not tied to any specific photo.",
    "",
    "IDENTITY AND CONTEXT:",
    "- When identity (name, role, or location) is provided, reference it naturally. Don't force it into every field if the photo doesn't depict the named subject.",
    "- When an existing caption (the \"seed\") is provided, it may contain real factual content (subjects, events, places) worth preserving. Strip promotional language from it but keep the facts. If the seed is purely promotional, generic, or just repeats the business name, ignore it.",
    "",
    "OUTPUT FORMAT:",
    "- Return ONLY the text for the field. No quotes wrapping it. No preamble (\"Here is the description:\"). No trailing ellipses.",
    "- If a photo is ambiguous or you cannot comply with the rules, return a short factual fallback like \"Photo of [identity]\" rather than promotional filler.",
    "- Do not include Yext token literals like [[name]] or [[website.url]] in the output — write plain text."
  ].join("\n");

  // ===========================
  // STATE
  // ===========================
  const STATE_KEY = Symbol.for("__mlYextMediaState_v23");
  const STATE = (window[STATE_KEY] = window[STATE_KEY] || {
    altState: new Map(),
    lastRunKeys: [],
    entityTitle: "",
    tokenMap: null,
    tokenMapBuiltAt: 0,
    aborted: false,
    running: false
  });

  // Keyed by stable item key (derived from textarea name), not array index.
  const SELECTED_KEYS = new Set();

  // ===========================
  // UTILITIES
  // ===========================

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function normalizeSpaces(s) {
    return (s || "")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .replace(/\s+/g, " ")
      .trim();
  }

  function sanitizeText(raw) {
    let s = (raw || "").toString();

    s = s.replace(/\r\n/g, " ").replace(/[\r\n]+/g, " ");
    s = s.replace(/[\u2014\u2013\u2212]/g, "-");

    // Drop ASCII double quote, smart double quotes (U+201C–U+201F), and backtick.
    s = s.replace(/[\u201C\u201D\u201E\u201F"`]/g, "");
    s = s.replace(/\u2018/g, "");   // left smart single quote → drop
    s = s.replace(/\u2019/g, "'");  // right smart single → straight apostrophe

    s = s.replace(/[\/\\]/g, " ");
    s = s.replace(/#/g, "");

    s = s.replace(/[\u{1F000}-\u{1FAFF}]/gu, "");
    s = s.replace(/[\u{2600}-\u{27BF}]/gu, "");
    s = s.replace(/[\u{FE00}-\u{FE0F}]/gu, "");
    s = s.replace(/[\u{1F900}-\u{1F9FF}]/gu, "");

    s = normalizeSpaces(s);

    if (s.length > CONFIG.maxChars) {
      s = s.slice(0, CONFIG.maxChars).trim();
    }

    return s;
  }

  function scrubPII(s) {
    if (!s) return "";

    const patterns = {
      email: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
      phone: /\b(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b/g,
      zip: /\b\d{5}(?:-\d{4})?\b/g,
      streetNum: /\b\d+\s+(?:N|S|E|W|North|South|East|West)?\s*[A-Z][a-z]+(?:\s+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Lane|Ln|Drive|Dr|Court|Ct|Circle|Cir|Way|Parkway|Pkwy|Place|Pl))\b/gi,
      poBox: /\bP\.?O\.?\s+Box\s+\d+\b/gi
    };

    let result = s;
    for (const pattern of Object.values(patterns)) {
      result = result.replace(pattern, " ");
    }

    return result;
  }

  function enforceFieldSpec(text) {
    if (!text) return "";
    let s = sanitizeText(text);
    s = scrubPII(s);
    s = normalizeSpaces(s);
    return s;
  }

  function validateFieldSpec(text) {
    const t = text || "";
    const errors = [];

    if (t.length < CONFIG.minChars) errors.push("too_short");
    if (t.length > CONFIG.maxChars) errors.push("too_long");
    if (/\n/.test(t)) errors.push("newlines");
    if (/  +/.test(t)) errors.push("double_spaces");
    if (/[\u201C\u201D\u201E\u201F"`\/\\#]/.test(t)) errors.push("forbidden_chars");
    if (/[\u2014\u2013]/.test(t)) errors.push("em_dash");
    if (/[\u{1F000}-\u{1FAFF}]/gu.test(t)) errors.push("emoji");

    if (/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/.test(t)) errors.push("email");
    if (/\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b/.test(t)) errors.push("phone");

    for (const token of FORBIDDEN_OUTPUT_TOKENS) {
      if (t.includes(token)) {
        errors.push("forbidden_token");
        break;
      }
    }

    return { ok: errors.length === 0, errors };
  }

  // Reuses a single toast node to avoid the flicker of remove+recreate.
  function toast(msg, durationMs = 4000) {
    let t = document.getElementById("ml-yext-toast");
    if (!t) {
      t = document.createElement("div");
      t.id = "ml-yext-toast";
      t.style.cssText = [
        "position:fixed", "bottom:24px", "right:24px", "z-index:2147483647",
        "background:#2c3e50", "color:#ecf0f1", "padding:14px 20px",
        "border-radius:8px", "font-family:system-ui,sans-serif",
        "font-size:13px", "box-shadow:0 4px 12px rgba(0,0,0,0.3)",
        "max-width:400px", "word-wrap:break-word", "line-height:1.4"
      ].join(";");
      document.body.appendChild(t);
    }

    t.textContent = msg;
    clearTimeout(t._hideTimer);
    t._hideTimer = setTimeout(() => {
      if (t.parentNode) t.parentNode.removeChild(t);
    }, durationMs);
  }

  // ===========================
  // GM FETCH WRAPPERS
  // ===========================

  function gmFetch(details) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        ...details,
        onload: (response) => resolve(response),
        onerror: (error) => reject(error),
        ontimeout: () => reject(new Error("Request timeout"))
      });
    });
  }

  async function fetchWithRetry(url, options = {}, maxRetries = CONFIG.apiRetryAttempts) {
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const response = await gmFetch({
          method: options.method || "GET",
          url: url,
          headers: options.headers || {},
          data: options.body || undefined,
          responseType: options.responseType
        });

        if (response.status === 429) {
          const delay = CONFIG.apiRetryBaseDelay * Math.pow(2, attempt);
          console.warn(`Rate limited. Retrying in ${delay}ms... (attempt ${attempt + 1}/${maxRetries})`);
          await sleep(delay);
          continue;
        }

        if (response.status >= 500 && response.status < 600) {
          if (attempt < maxRetries - 1) {
            const delay = CONFIG.apiRetryBaseDelay * Math.pow(2, attempt);
            console.warn(`Server error ${response.status}. Retrying in ${delay}ms...`);
            await sleep(delay);
            continue;
          }
        }

        return response;
      } catch (error) {
        if (attempt === maxRetries - 1) {
          throw new Error(`Fetch failed after ${maxRetries} attempts: ${error.message}`);
        }
        const delay = CONFIG.apiRetryBaseDelay * Math.pow(2, attempt);
        console.warn(`Network error. Retrying in ${delay}ms...`, error);
        await sleep(delay);
      }
    }

    throw new Error(`Failed to fetch after ${maxRetries} attempts`);
  }

  async function callGemini(prompt, imageDataUrl = null, systemPrompt = null) {
    const apiKey = CONFIG.getGeminiApiKey();
    if (!apiKey) {
      throw new Error("Gemini API key not set. Run GM_setValue('geminiApiKey', '<your-key>') in the console.");
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.geminiModel}:generateContent?key=${apiKey}`;

    // Gemini has no prompt cache like Claude does, so we inline the system prompt.
    const combinedPrompt = systemPrompt ? `${systemPrompt}\n\n${prompt}` : prompt;

    const parts = [];
    if (combinedPrompt) parts.push({ text: combinedPrompt });

    if (imageDataUrl) {
      if (imageDataUrl.length > CONFIG.maxDataUrlChars) {
        throw new Error(`Image data URL too large: ${imageDataUrl.length} chars (max: ${CONFIG.maxDataUrlChars})`);
      }

      const match = imageDataUrl.match(/^data:image\/(\w+);base64,(.+)$/);
      if (!match) throw new Error("Invalid image data URL format");

      parts.push({
        inline_data: {
          mime_type: `image/${match[1]}`,
          data: match[2]
        }
      });
    }

    const body = JSON.stringify({ contents: [{ parts }] });

    const resp = await fetchWithRetry(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body
    });

    if (resp.status < 200 || resp.status >= 300) {
      throw new Error(`Gemini API error ${resp.status}: ${resp.responseText || resp.statusText}`);
    }

    const data = JSON.parse(resp.responseText);

    if (!data.candidates || !Array.isArray(data.candidates) || data.candidates.length === 0) {
      throw new Error("Invalid API response: no candidates returned");
    }

    const candidate = data.candidates[0];
    if (!candidate.content || !candidate.content.parts || candidate.content.parts.length === 0) {
      throw new Error("Invalid API response: no content parts");
    }

    const text = candidate.content.parts[0].text;
    if (typeof text !== "string") {
      throw new Error("Invalid API response: text is not a string");
    }

    return text.trim();
  }

  async function callLocalCaption(imageDataUrl) {
    if (!imageDataUrl) throw new Error("No image data URL provided");

    const resp = await fetchWithRetry(CONFIG.localCaptionUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image: imageDataUrl })
    });

    if (resp.status < 200 || resp.status >= 300) {
      throw new Error(`Local caption service error ${resp.status}: ${resp.responseText || resp.statusText}`);
    }

    const data = JSON.parse(resp.responseText);
    if (!data.caption || typeof data.caption !== "string") {
      throw new Error("Invalid response from caption service");
    }
    return data.caption.trim();
  }

  // Anthropic Messages API. System prompt is passed as a cacheable block — the
  // `cache_control: ephemeral` marker tells Anthropic to serve the system prefix
  // from cache on repeat calls at ~10% of normal input cost. The marker is a
  // silent no-op if the prefix is below the model's caching minimum.
  async function callClaude(userPrompt, imageDataUrl = null, systemPrompt = null) {
    const apiKey = CONFIG.getAnthropicApiKey();
    if (!apiKey) {
      throw new Error("Anthropic API key not set. Run GM_setValue('anthropicApiKey', '<your-key>') in the console.");
    }

    const content = [];
    if (imageDataUrl) {
      if (imageDataUrl.length > CONFIG.maxDataUrlChars) {
        throw new Error(`Image data URL too large: ${imageDataUrl.length} chars (max: ${CONFIG.maxDataUrlChars})`);
      }
      const match = imageDataUrl.match(/^data:image\/(\w+);base64,(.+)$/);
      if (!match) throw new Error("Invalid image data URL format");
      content.push({
        type: "image",
        source: { type: "base64", media_type: `image/${match[1]}`, data: match[2] }
      });
    }
    content.push({ type: "text", text: userPrompt });

    const body = {
      model: CONFIG.claudeModel,
      max_tokens: 300,
      messages: [{ role: "user", content }]
    };

    if (systemPrompt) {
      body.system = [
        { type: "text", text: systemPrompt, cache_control: { type: "ephemeral" } }
      ];
    }

    const resp = await fetchWithRetry("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    if (resp.status < 200 || resp.status >= 300) {
      throw new Error(`Claude API error ${resp.status}: ${resp.responseText || resp.statusText}`);
    }

    const data = JSON.parse(resp.responseText);

    if (!data.content || !Array.isArray(data.content) || data.content.length === 0) {
      throw new Error("Invalid Claude response: no content array");
    }

    const textBlock = data.content.find((b) => b.type === "text");
    if (!textBlock || typeof textBlock.text !== "string") {
      throw new Error("Invalid Claude response: no text block");
    }

    return textBlock.text.trim();
  }

  // Pick the provider for a given field. "hybrid" uses Claude for the nuanced
  // fields (desc/details/business) and Gemini for alt (short, literal).
  function chooseProvider(field) {
    if (CONFIG.provider === "gemini") return "gemini";
    if (CONFIG.provider === "claude") return "claude";
    return field === "alt" ? "gemini" : "claude";
  }

  // Unified dispatcher used by all generators. Routes by provider config + field.
  async function callLLM({ userPrompt, imageDataUrl = null, field = "description", systemPrompt = SYSTEM_PROMPT }) {
    const provider = chooseProvider(field);
    if (provider === "claude") {
      return callClaude(userPrompt, imageDataUrl, systemPrompt);
    }
    return callGemini(userPrompt, imageDataUrl, systemPrompt);
  }

  // Build a compact, per-call user prompt. The bulk of the guidance lives in
  // SYSTEM_PROMPT (cacheable on Claude); the user prompt only carries the
  // dynamic details that change per item.
  function buildUserPrompt({ field, identity, seed = "", extras = {} }) {
    const parts = [];
    parts.push(`Field: ${field}`);
    parts.push(`Identity: ${identity || "(none provided)"}`);
    if (extras.location) parts.push(`Location: ${extras.location}`);
    if (seed) parts.push(`Existing caption (strip promotional language, preserve factual subjects/places/events): "${seed}"`);

    if (field === "description") {
      parts.push("Task: Write a factual description of what's specifically in this photo (5 to 249 chars).");
    } else if (field === "details") {
      parts.push("Task: Write complementary details (context, mood, or a fact the description doesn't cover). Must differ from the description (5 to 249 chars).");
    } else if (field === "alt") {
      parts.push("Task: Write short literal alt text describing what is visible in the frame (prefer under 50 chars).");
    } else if (field === "business_description") {
      parts.push("Task: Write a single evergreen business description (80 to 200 chars). Not tied to any one photo. No promotional language.");
    }

    return parts.join("\n");
  }

  // ===========================
  // IMAGE PROCESSING
  // ===========================

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("FileReader failed"));
      reader.readAsDataURL(blob);
    });
  }

  // Fetch through GM_xmlhttpRequest to bypass the canvas CORS taint that would occur
  // if we loaded a cross-origin <img> and called toDataURL.
  async function fetchImageAsDataUrl(src) {
    const resp = await gmFetch({
      method: "GET",
      url: src,
      responseType: "blob"
    });
    if (resp.status < 200 || resp.status >= 300) {
      throw new Error(`Image fetch failed: ${resp.status}`);
    }
    return await blobToDataUrl(resp.response);
  }

  function resizeImage(dataUrl, maxDim) {
    return new Promise((resolve, reject) => {
      const img = new Image();

      img.onload = () => {
        try {
          let w = img.width;
          let h = img.height;

          if (w <= maxDim && h <= maxDim) {
            resolve(dataUrl);
            return;
          }

          if (w > h) {
            h = Math.round((h * maxDim) / w);
            w = maxDim;
          } else {
            w = Math.round((w * maxDim) / h);
            h = maxDim;
          }

          const canvas = document.createElement("canvas");
          canvas.width = w;
          canvas.height = h;

          const ctx = canvas.getContext("2d");
          if (!ctx) {
            reject(new Error("Failed to get canvas context"));
            return;
          }

          ctx.drawImage(img, 0, 0, w, h);
          resolve(canvas.toDataURL("image/jpeg", CONFIG.jpegQuality));
        } catch (error) {
          reject(new Error(`Image resize failed: ${error.message}`));
        }
      };

      img.onerror = () => reject(new Error("Failed to load image for resizing"));
      img.src = dataUrl;
    });
  }

  async function getImageDataUrl(item) {
    try {
      if (!item || !item.img) return null;

      const imgEl = item.img;
      const src = imgEl.src || imgEl.getAttribute("src");
      if (!src) return null;

      if (src.startsWith("data:")) {
        if (src.length > CONFIG.maxDataUrlChars) {
          return await resizeImage(src, CONFIG.maxDim);
        }
        return src;
      }

      let dataUrl = await fetchImageAsDataUrl(src);
      if (dataUrl.length > CONFIG.maxDataUrlChars) {
        dataUrl = await resizeImage(dataUrl, CONFIG.maxDim);
      }
      return dataUrl;
    } catch (error) {
      console.error("getImageDataUrl failed:", error);
      return null;
    }
  }

  // Per-run cache so each item's image is fetched and encoded at most once,
  // even though three generators (desc/details/alt) consume it.
  function createImageCache() {
    const cache = new Map();
    return {
      get(item) {
        const key = getItemKey(item);
        if (cache.has(key)) return cache.get(key);
        const p = getImageDataUrl(item);
        cache.set(key, p);
        return p;
      },
      clear() { cache.clear(); }
    };
  }

  // ===========================
  // DOM UTILITIES
  // ===========================

  // Derive a stable identifier from the textarea's name attribute (e.g. "c_photos.0.description" → "c_photos.0").
  // Falls back to positional index if names are unavailable.
  function getItemKey(item) {
    const n = item.desc?.name || item.details?.name || item.alt?.name || "";
    const base = n.replace(/\.(description|details|alternateText)$/, "");
    return base || `item_${item.index}`;
  }

  function findMediaItems() {
    const sel = CONFIG.selectors;

    // Card-anchored discovery: each photo card contains its own desc/details/alt.
    // Walk up from each desc until we find the narrowest ancestor that contains
    // exactly 1 desc plus details and alt — that ancestor is the card, and its
    // details/alt are guaranteed to belong to this desc (not a sibling card's).
    const descSelector = `${sel.description}, ${sel.fallbackDescription}`;
    const detailSelector = `${sel.details}, ${sel.fallbackDetails}`;
    const altSelector = `${sel.alternateText}, ${sel.fallbackAlt}`;

    const descs = Array.from(document.querySelectorAll(descSelector));
    const seen = new WeakSet();
    const items = [];
    const MAX_DEPTH = 12;

    descs.forEach((desc, index) => {
      if (seen.has(desc)) return;
      seen.add(desc);

      let card = desc.parentElement;
      let depth = 0;
      let found = null;

      while (card && depth < MAX_DEPTH) {
        const descsHere = card.querySelectorAll(descSelector);
        if (descsHere.length === 1) {
          const detailsHere = card.querySelectorAll(detailSelector);
          const altsHere = card.querySelectorAll(altSelector);
          if (detailsHere.length >= 1 && altsHere.length >= 1) {
            found = {
              details: detailsHere[0],
              alt: altsHere[0],
              clickthrough: card.querySelector(sel.clickthroughUrl),
              img: card.querySelector("img")
            };
            break;
          }
        }
        card = card.parentElement;
        depth++;
      }

      items.push({
        index,
        desc,
        details: found?.details || null,
        alt: found?.alt || null,
        clickthrough: found?.clickthrough || null,
        img: found?.img || null
      });
    });

    // Fallback: if no descs found at all, try positional pairing on fallback selectors.
    if (items.length === 0) {
      const details = Array.from(document.querySelectorAll(sel.fallbackDetails));
      const alts = Array.from(document.querySelectorAll(sel.fallbackAlt));
      const clicks = Array.from(document.querySelectorAll(sel.clickthroughUrl));
      const n = Math.max(details.length, alts.length, clicks.length);
      for (let i = 0; i < n; i++) {
        items.push({
          index: i,
          desc: null,
          details: details[i] || null,
          alt: alts[i] || null,
          clickthrough: clicks[i] || null,
          img: null
        });
      }
    }

    return items;
  }

  function getEntityTitleText() {
    return getNameFromNameField() || getNameFromBreadcrumb() || getNameFromPageTitle() || "";
  }

  function getNameFromNameField() {
    const nameInput =
      document.querySelector('input[aria-label="name"]') ||
      document.querySelector('input[name*="name"]');
    return nameInput ? nameInput.value.trim() : "";
  }

  function getNameFromBreadcrumb() {
    const breadcrumbs = Array.from(document.querySelectorAll('[class*="breadcrumb"] a, nav a'));
    if (breadcrumbs.length > 0) return breadcrumbs[breadcrumbs.length - 1].textContent.trim();
    return "";
  }

  function getNameFromPageTitle() {
    const parts = (document.title || "").split(/[|\-—]/);
    return parts.length > 0 ? parts[0].trim() : "";
  }

  // Populate as many tokens as possible by probing for inputs whose aria-label or name
  // matches the token's key (full path or leaf). 5-second cache keeps it cheap.
  function getTokenMap(entityTitle) {
    const now = Date.now();
    if (STATE.tokenMap && (now - STATE.tokenMapBuiltAt) < 5000) {
      return STATE.tokenMap;
    }

    const map = {};
    map["[[name]]"] = entityTitle || "";

    const fieldKey = (token) => token.replace(/^\[\[|\]\]$/g, "");

    for (const token of EMBEDDED_TOKENS) {
      if (map[token]) continue;
      const key = fieldKey(token);
      const leaf = key.split(".").pop();

      const el =
        document.querySelector(`input[aria-label="${key}"]`) ||
        document.querySelector(`input[aria-label="${leaf}"]`) ||
        document.querySelector(`input[aria-label*="${key}"]`) ||
        document.querySelector(`input[aria-label*="${leaf}"]`) ||
        document.querySelector(`input[name="${key}"]`) ||
        document.querySelector(`input[name*="${key}"]`);

      if (el && el.value) {
        map[token] = el.value.trim();
      }
    }

    STATE.tokenMap = map;
    STATE.tokenMapBuiltAt = now;
    return map;
  }

  function extractNameRoleFromTitle(title) {
    if (!title) return { name: "", role: "" };

    const separators = [" - ", ", ", " – "];
    for (const sep of separators) {
      if (title.includes(sep)) {
        const parts = title.split(sep);
        return {
          name: parts[0].trim(),
          role: parts.slice(1).join(sep).trim()
        };
      }
    }

    return { name: title, role: "" };
  }

  // ===========================
  // TEXT ANALYSIS
  // ===========================

  function isPromotional(text) {
    return !!text && PROMO_REGEX.test(text);
  }

  function stripUrlTokens(text) {
    if (!text) return "";
    let result = text;
    for (const token of FORBIDDEN_OUTPUT_TOKENS) {
      result = result.replace(new RegExp(escapeRegex(token), "g"), "");
    }
    return normalizeSpaces(result);
  }

  function descIsJustEntityName(desc, entityTitle, rawDesc) {
    if (!desc || !entityTitle) return false;

    const descClean = stripUrlTokens(desc).toLowerCase().trim();
    const titleClean = entityTitle.toLowerCase().trim();

    if (descClean === titleClean) return true;
    if (descClean === `photo of ${titleClean}`) return true;
    if ((rawDesc || "").trim() === "[[name]]") return true;

    return false;
  }

  function revertLiteralToEmbedded(text, entityTitle) {
    if (!text) return "";
    const tokenMap = getTokenMap(entityTitle);
    let result = text;

    // Composite literals (e.g. "City, ST 12345") — replace these before per-token
    // passes so the combined form is preserved instead of partially tokenized.
    const line1 = tokenMap["[[address.line1]]"];
    const city = tokenMap["[[address.city]]"];
    const region = tokenMap["[[address.region]]"];
    const zip = tokenMap["[[address.postalCode]]"];

    const composites = [];
    if (line1 && city && region && zip) {
      composites.push([
        `${line1}, ${city}, ${region} ${zip}`,
        "[[address.line1]], [[address.city]], [[address.region]] [[address.postalCode]]"
      ]);
    }
    if (city && region && zip) {
      composites.push([
        `${city}, ${region} ${zip}`,
        "[[address.city]], [[address.region]] [[address.postalCode]]"
      ]);
    }
    if (city && region) {
      composites.push([`${city}, ${region}`, "[[address.city]], [[address.region]]"]);
    }

    composites.sort((a, b) => b[0].length - a[0].length);
    for (const [literal, replacement] of composites) {
      result = result.replace(new RegExp(escapeRegex(literal), "g"), replacement);
    }

    const entries = Object.entries(tokenMap).sort((a, b) => b[1].length - a[1].length);
    for (const [token, value] of entries) {
      if (value && result.includes(value)) {
        result = result.replace(new RegExp(escapeRegex(value), "g"), token);
      }
    }

    return result;
  }

  // ===========================
  // CONTENT GENERATION
  // ===========================

  function deriveIdentityForAlt(item, entityTitle, tokenMap) {
    const { name, role } = extractNameRoleFromTitle(entityTitle);
    if (name && role) return `${name}, ${role}`;
    if (entityTitle) return entityTitle;
    return tokenMap?.["[[name]]"] || "";
  }

  function buildDescriptionSuggestion(item, entityTitle) {
    const { name } = extractNameRoleFromTitle(entityTitle);
    if (name) return `Photo of ${name}`;
    if (entityTitle) return `Photo of ${entityTitle}`;
    return "Photo";
  }

  function buildDetailsSuggestion(item, entityTitle) {
    return buildDescriptionSuggestion(item, entityTitle);
  }

  // All three generators accept a pre-fetched imageDataUrl so a single image fetch
  // covers all three calls. `seed` is the original caption text (optional).
  async function generateAlt(item, entityTitle, tokenMap, imageDataUrl, seed = "") {
    const itemKey = getItemKey(item);
    const state = STATE.altState.get(itemKey) || { altVariantIndex: 0, cachedCaption: null, visionFailed: false };

    if (CONFIG.altMode === "vision" && !state.visionFailed && imageDataUrl) {
      try {
        const identity = deriveIdentityForAlt(item, entityTitle, tokenMap);
        const userPrompt = buildUserPrompt({ field: "alt", identity, seed });
        const altText = await callLLM({ userPrompt, imageDataUrl, field: "alt" });
        return enforceFieldSpec(altText);
      } catch (error) {
        console.error(`Vision alt generation failed for item ${item.index}:`, error);
        state.visionFailed = true;
        STATE.altState.set(itemKey, state);
      }
    }

    const identity = deriveIdentityForAlt(item, entityTitle, tokenMap);
    return enforceFieldSpec(`Photo of ${identity}`);
  }

  async function generateDescription(item, entityTitle, tokenMap, imageDataUrl, seed = "") {
    try {
      const identity = deriveIdentityForAlt(item, entityTitle, tokenMap);
      const userPrompt = buildUserPrompt({ field: "description", identity, seed });
      const desc = await callLLM({ userPrompt, imageDataUrl: imageDataUrl || null, field: "description" });
      return enforceFieldSpec(desc);
    } catch (error) {
      console.error(`Description generation failed for item ${item.index}:`, error);
      return buildDescriptionSuggestion(item, entityTitle);
    }
  }

  async function generateDetails(item, entityTitle, tokenMap, imageDataUrl, seed = "") {
    try {
      const identity = deriveIdentityForAlt(item, entityTitle, tokenMap);
      const userPrompt = buildUserPrompt({ field: "details", identity, seed });
      const details = await callLLM({ userPrompt, imageDataUrl: imageDataUrl || null, field: "details" });
      return enforceFieldSpec(details);
    } catch (error) {
      console.error(`Details generation failed for item ${item.index}:`, error);
      return buildDetailsSuggestion(item, entityTitle);
    }
  }

  // Entity-level evergreen description used when the shared-description mode is on.
  async function generateBusinessDescription(entityTitle, tokenMap) {
    const identity = deriveIdentityForAlt(null, entityTitle, tokenMap) || entityTitle || "this business";
    const city = tokenMap["[[address.city]]"] || "";
    const region = tokenMap["[[address.region]]"] || "";
    const location = city && region ? `${city}, ${region}` : (city || "");
    const userPrompt = buildUserPrompt({
      field: "business_description",
      identity,
      extras: location ? { location } : {}
    });
    const desc = await callLLM({ userPrompt, field: "business_description" });
    return enforceFieldSpec(desc);
  }

  // ===========================
  // FIELD WRITE + FLOW CONTROL
  // ===========================

  function setFieldValue(el, value) {
    if (!el) return false;
    if (el.value === value) return false;
    // focus/blur wrap: some React-controlled inputs only commit state on blur,
    // so without this the written value can silently revert on navigation.
    try { el.focus(); } catch (_) {}
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    try { el.blur(); } catch (_) {}
    return true;
  }

  // Use the existing desc as the seed for all three generators. Skip very short
  // seeds (empty, single word, just the entity name) — they add no context and
  // can anchor Gemini to repeat the name.
  function deriveSeed(item) {
    const raw = (item?.desc?.value || "").trim();
    if (raw.length < 12) return "";
    if (/^\[\[.+?\]\]$/.test(raw)) return "";
    return raw;
  }

  function checkAborted() {
    if (STATE.aborted) throw new Error("Run cancelled by user.");
  }

  // Worker-pool parallelism: up to `limit` items in flight. Aborts cooperatively.
  async function runPool(items, limit, task) {
    const queue = items.slice();
    const active = Math.min(limit, queue.length);
    const workers = [];

    for (let i = 0; i < active; i++) {
      workers.push((async () => {
        while (queue.length > 0) {
          if (STATE.aborted) return;
          const item = queue.shift();
          await task(item);
        }
      })());
    }

    await Promise.all(workers);
  }

  // ===========================
  // MAIN OPERATIONS
  // ===========================

  async function runSanitize(items) {
    let processed = 0;
    let errors = 0;

    for (let i = 0; i < items.length; i++) {
      checkAborted();
      const item = items[i];

      try {
        if (item.desc) setFieldValue(item.desc, enforceFieldSpec(item.desc.value || ""));
        if (item.details) setFieldValue(item.details, enforceFieldSpec(item.details.value || ""));
        if (item.alt) setFieldValue(item.alt, enforceFieldSpec(item.alt.value || ""));
        processed++;
      } catch (error) {
        console.error(`Error sanitizing item ${item.index}:`, error);
        errors++;
      }

      if (i < items.length - 1) await sleep(CONFIG.pauseMsBetweenItems);
    }

    toast(`Sanitize complete: ${processed} processed, ${errors} errors`);
  }

  async function runOverwrite(items, entityTitle, tokenMap, options = {}) {
    const { previewMode = false, sharedDescription = false } = options;
    const imageCache = createImageCache();
    const previews = [];
    let processed = 0;
    let errors = 0;

    let sharedDesc = null;
    if (sharedDescription) {
      try {
        toast("Generating shared business description...");
        sharedDesc = await generateBusinessDescription(entityTitle, tokenMap);
      } catch (e) {
        toast(`Failed to generate shared description: ${e.message}`, 8000);
        return;
      }
    }

    toast(`Generating for ${items.length} item(s)${previewMode ? " (preview)" : ""}...`);

    await runPool(items, CONFIG.itemConcurrency, async (item) => {
      if (STATE.aborted) return;
      try {
        const seed = deriveSeed(item);
        const imageDataUrl = await imageCache.get(item);

        const descPromise = sharedDesc
          ? Promise.resolve(sharedDesc)
          : generateDescription(item, entityTitle, tokenMap, imageDataUrl, seed);

        const [desc, details, alt] = await Promise.all([
          descPromise,
          generateDetails(item, entityTitle, tokenMap, imageDataUrl, seed),
          generateAlt(item, entityTitle, tokenMap, imageDataUrl, seed)
        ]);

        if (previewMode) {
          previews.push({
            item,
            desc,
            details,
            alt,
            clickthrough: CONFIG.setClickthroughUrl && item.clickthrough ? CONFIG.clickthroughValue : null
          });
        } else {
          if (item.desc) setFieldValue(item.desc, desc);
          if (item.details) setFieldValue(item.details, details);
          if (item.alt) setFieldValue(item.alt, alt);
          if (CONFIG.setClickthroughUrl && item.clickthrough) {
            setFieldValue(item.clickthrough, CONFIG.clickthroughValue);
          }
        }

        processed++;
        toast(`Progress: ${processed}/${items.length} done`);
      } catch (error) {
        console.error(`Error processing item ${item.index}:`, error);
        errors++;
        toast(`Error on item ${item.index + 1}: ${error.message}`, 6000);
      }
    });

    if (previewMode && previews.length) showPreviewPanel(previews, { label: "Overwrite Preview" });

    toast(`Overwrite ${previewMode ? "preview" : "complete"}: ${processed} processed, ${errors} errors`, 5000);
  }

  async function runSuggest(items, entityTitle, tokenMap, options = {}) {
    const { previewMode = false, sharedDescription = false } = options;
    const imageCache = createImageCache();
    const previews = [];
    let processed = 0;
    let skipped = 0;
    let errors = 0;

    let sharedDesc = null;
    if (sharedDescription) {
      try {
        toast("Generating shared business description...");
        sharedDesc = await generateBusinessDescription(entityTitle, tokenMap);
      } catch (e) {
        toast(`Failed to generate shared description: ${e.message}`, 8000);
        return;
      }
    }

    toast(`Processing ${items.length} item(s) in suggest mode${previewMode ? " (preview)" : ""}...`);

    await runPool(items, CONFIG.itemConcurrency, async (item) => {
      if (STATE.aborted) return;
      try {
        const needs = { desc: false, details: false, alt: false, clickthrough: false };

        if (item.desc) {
          const val = item.desc.value || "";
          const reverted = revertLiteralToEmbedded(val, entityTitle);
          const stripped = stripUrlTokens(reverted);
          const validation = validateFieldSpec(val);
          needs.desc = !val || !validation.ok || descIsJustEntityName(stripped, entityTitle, val) || isPromotional(val);
        }

        if (item.details) {
          const val = item.details.value || "";
          const validation = validateFieldSpec(val);
          needs.details = !val || !validation.ok || isPromotional(val);
        }

        if (item.alt) {
          const val = item.alt.value || "";
          const validation = validateFieldSpec(val);
          needs.alt = !val || !validation.ok || isPromotional(val);
        }

        if (CONFIG.setClickthroughUrl && item.clickthrough && !item.clickthrough.value) {
          needs.clickthrough = true;
        }

        if (!needs.desc && !needs.details && !needs.alt && !needs.clickthrough) {
          skipped++;
          return;
        }

        const seed = deriveSeed(item);

        // Only fetch the image if at least one image-based generation needs it.
        // Shared-desc path skips the image fetch for desc but details/alt still need it.
        const imageNeeded = (needs.desc && !sharedDesc) || needs.details || needs.alt;
        const imageDataUrl = imageNeeded ? await imageCache.get(item) : null;

        const descPromise = needs.desc
          ? (sharedDesc ? Promise.resolve(sharedDesc) : generateDescription(item, entityTitle, tokenMap, imageDataUrl, seed))
          : Promise.resolve(null);

        const [descVal, detailsVal, altVal] = await Promise.all([
          descPromise,
          needs.details ? generateDetails(item, entityTitle, tokenMap, imageDataUrl, seed) : Promise.resolve(null),
          needs.alt ? generateAlt(item, entityTitle, tokenMap, imageDataUrl, seed) : Promise.resolve(null)
        ]);

        if (previewMode) {
          previews.push({
            item,
            desc: needs.desc ? descVal : null,
            details: needs.details ? detailsVal : null,
            alt: needs.alt ? altVal : null,
            clickthrough: needs.clickthrough ? CONFIG.clickthroughValue : null
          });
        } else {
          if (needs.desc) setFieldValue(item.desc, descVal);
          if (needs.details) setFieldValue(item.details, detailsVal);
          if (needs.alt) setFieldValue(item.alt, altVal);
          if (needs.clickthrough) setFieldValue(item.clickthrough, CONFIG.clickthroughValue);
        }

        processed++;
        toast(`Progress: ${processed} updated, ${skipped} skipped, ${errors} errors`);
      } catch (error) {
        console.error(`Error processing item ${item.index}:`, error);
        errors++;
        toast(`Error on item ${item.index + 1}: ${error.message}`, 6000);
      }
    });

    if (previewMode && previews.length) showPreviewPanel(previews, { label: "Suggest Preview" });

    toast(`Suggest ${previewMode ? "preview" : "complete"}: ${processed} updated, ${skipped} skipped, ${errors} errors`, 5000);
  }

  // ===========================
  // UI — SELECTION CHECKBOXES
  // ===========================

  function addSelectionCheckboxes() {
    const items = findMediaItems();

    items.forEach((item) => {
      if (!item.desc) return;

      const container = item.desc.closest('[class*="card"]') || item.desc.closest("div");
      if (!container) return;
      if (container.querySelector(".ml-yext-select-checkbox")) return;

      const key = getItemKey(item);
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "ml-yext-select-checkbox";
      checkbox.dataset.itemKey = key;
      checkbox.style.cssText = "position:absolute;top:8px;left:8px;z-index:1000;width:20px;height:20px;cursor:pointer;";
      checkbox.checked = SELECTED_KEYS.has(key);

      checkbox.addEventListener("change", () => {
        if (checkbox.checked) SELECTED_KEYS.add(key);
        else SELECTED_KEYS.delete(key);
      });

      container.style.position = "relative";
      container.appendChild(checkbox);
    });
  }

  function removeSelectionCheckboxes() {
    document.querySelectorAll(".ml-yext-select-checkbox").forEach((cb) => cb.remove());
    SELECTED_KEYS.clear();
  }

  // ===========================
  // UI — PREVIEW PANEL
  // ===========================

  function escapeHtml(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function showPreviewPanel(previews, { label = "Preview" } = {}) {
    const existing = document.getElementById("ml-yext-preview-panel");
    if (existing) existing.remove();

    const panel = document.createElement("div");
    panel.id = "ml-yext-preview-panel";
    panel.style.cssText = [
      "position:fixed", "top:16px", "right:16px", "z-index:2147483647",
      "width:520px", "max-height:85vh", "overflow:auto",
      "background:#fff", "color:#222", "font-family:system-ui,sans-serif",
      "font-size:12px", "line-height:1.4", "padding:14px 16px",
      "border-radius:10px", "border:1px solid #ccc",
      "box-shadow:0 6px 22px rgba(0,0,0,0.2)"
    ].join(";");

    const header = document.createElement("div");
    header.style.cssText = "display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;";
    const title = document.createElement("h3");
    title.textContent = `${label} (${previews.length} items)`;
    title.style.cssText = "margin:0;font-size:14px;";
    header.appendChild(title);
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "✕";
    closeBtn.type = "button";
    closeBtn.style.cssText = "background:none;border:none;cursor:pointer;font-size:16px;";
    closeBtn.addEventListener("click", () => panel.remove());
    header.appendChild(closeBtn);
    panel.appendChild(header);

    const selectAllRow = document.createElement("div");
    selectAllRow.style.cssText = "margin-bottom:8px;";
    const selectAll = document.createElement("input");
    selectAll.type = "checkbox";
    selectAll.checked = true;
    selectAll.id = "ml-yext-preview-selectall";
    const selectAllLabel = document.createElement("label");
    selectAllLabel.htmlFor = "ml-yext-preview-selectall";
    selectAllLabel.textContent = " Select all";
    selectAllLabel.style.cssText = "cursor:pointer;user-select:none;";
    selectAllRow.appendChild(selectAll);
    selectAllRow.appendChild(selectAllLabel);
    panel.appendChild(selectAllRow);

    const rowsContainer = document.createElement("div");
    panel.appendChild(rowsContainer);

    const checkboxes = [];

    previews.forEach((p) => {
      const row = document.createElement("div");
      row.style.cssText = "border-top:1px solid #eee;padding:8px 0;";

      const head = document.createElement("div");
      head.style.cssText = "display:flex;gap:8px;align-items:center;margin-bottom:4px;font-weight:600;";

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = true;
      checkboxes.push(cb);
      head.appendChild(cb);

      const lbl = document.createElement("span");
      lbl.textContent = `Item ${p.item.index + 1}`;
      head.appendChild(lbl);

      row.appendChild(head);

      const addField = (name, value, oldValue) => {
        if (value == null) return;
        const f = document.createElement("div");
        f.style.cssText = "margin-left:22px;margin-bottom:4px;";
        const oldStr = oldValue != null && oldValue !== "" ? `<div style="color:#aaa;text-decoration:line-through;">${escapeHtml(oldValue)}</div>` : "";
        f.innerHTML = `<span style="color:#888;font-weight:600;">${name}:</span> ${oldStr}<div>${escapeHtml(value)}</div>`;
        row.appendChild(f);
      };

      addField("desc", p.desc, p.item.desc?.value);
      addField("details", p.details, p.item.details?.value);
      addField("alt", p.alt, p.item.alt?.value);
      addField("clickthrough", p.clickthrough, p.item.clickthrough?.value);

      rowsContainer.appendChild(row);
    });

    selectAll.addEventListener("change", () => {
      checkboxes.forEach((cb) => { cb.checked = selectAll.checked; });
    });

    const btnRow = document.createElement("div");
    btnRow.style.cssText = "margin-top:12px;display:flex;gap:8px;";

    const applyBtn = document.createElement("button");
    applyBtn.textContent = "Apply Selected";
    applyBtn.type = "button";
    applyBtn.style.cssText = "flex:1;padding:8px 12px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:600;";
    applyBtn.addEventListener("click", () => {
      let applied = 0;
      previews.forEach((p, i) => {
        if (!checkboxes[i].checked) return;
        if (p.desc != null && p.item.desc) setFieldValue(p.item.desc, p.desc);
        if (p.details != null && p.item.details) setFieldValue(p.item.details, p.details);
        if (p.alt != null && p.item.alt) setFieldValue(p.item.alt, p.alt);
        if (p.clickthrough != null && p.item.clickthrough) setFieldValue(p.item.clickthrough, p.clickthrough);
        applied++;
      });
      panel.remove();
      toast(`Applied ${applied} item(s)`);
    });
    btnRow.appendChild(applyBtn);

    const discardBtn = document.createElement("button");
    discardBtn.textContent = "Discard";
    discardBtn.type = "button";
    discardBtn.style.cssText = "padding:8px 12px;background:#f0f0f0;border:1px solid #ccc;border-radius:6px;cursor:pointer;";
    discardBtn.addEventListener("click", () => panel.remove());
    btnRow.appendChild(discardBtn);

    panel.appendChild(btnRow);
    document.body.appendChild(panel);
  }

  // ===========================
  // UI — CONTROL PANEL
  // ===========================

  function mountControls() {
    const existing = document.getElementById("ml-yext-controls");
    if (existing) existing.remove();

    const wrap = document.createElement("div");
    wrap.id = "ml-yext-controls";
    wrap.style.cssText = [
      "position:fixed", "bottom:24px", "left:24px", "z-index:2147483647",
      "background:#ffffff", "padding:16px 20px", "border-radius:12px",
      "box-shadow:0 4px 16px rgba(0,0,0,0.15)", "font-family:system-ui,sans-serif",
      "font-size:13px", "min-width:340px", "border:1px solid #e0e0e0"
    ].join(";");

    const modeLabel = document.createElement("label");
    modeLabel.textContent = "Mode: ";
    modeLabel.style.cssText = "font-weight:600;margin-right:8px;color:#333;";

    const select = document.createElement("select");
    select.style.cssText = "padding:6px 10px;border:1px solid #ccc;border-radius:6px;margin-right:12px;font-size:13px;";
    ["suggest", "overwrite", "sanitize"].forEach((mode) => {
      const opt = document.createElement("option");
      opt.value = mode;
      opt.textContent = mode.charAt(0).toUpperCase() + mode.slice(1);
      if (mode === CONFIG.defaultMode) opt.selected = true;
      select.appendChild(opt);
    });

    const scopeLabel = document.createElement("label");
    scopeLabel.textContent = "Scope: ";
    scopeLabel.style.cssText = "font-weight:600;margin-right:8px;margin-left:4px;color:#333;";

    const scopeSelect = document.createElement("select");
    scopeSelect.style.cssText = "padding:6px 10px;border:1px solid #ccc;border-radius:6px;font-size:13px;";
    ["all", "selected", "single"].forEach((scope) => {
      const opt = document.createElement("option");
      opt.value = scope;
      opt.textContent = scope.charAt(0).toUpperCase() + scope.slice(1);
      if (scope === "all") opt.selected = true;
      scopeSelect.appendChild(opt);
    });

    const singleSelect = document.createElement("select");
    singleSelect.style.cssText = "padding:6px 10px;border:1px solid #ccc;border-radius:6px;margin-left:8px;display:none;font-size:13px;";

    const updateSingleSelect = () => {
      singleSelect.innerHTML = "";
      const items = findMediaItems();
      items.forEach((item) => {
        const opt = document.createElement("option");
        opt.value = getItemKey(item);
        opt.textContent = `Item ${item.index + 1}`;
        singleSelect.appendChild(opt);
      });
    };

    scopeSelect.addEventListener("change", () => {
      const scope = scopeSelect.value;
      if (scope === "single") {
        singleSelect.style.display = "inline-block";
        updateSingleSelect();
        removeSelectionCheckboxes();
      } else if (scope === "selected") {
        singleSelect.style.display = "none";
        addSelectionCheckboxes();
      } else {
        singleSelect.style.display = "none";
        removeSelectionCheckboxes();
      }
    });

    const providerLabel = document.createElement("label");
    providerLabel.textContent = "Provider: ";
    providerLabel.style.cssText = "font-weight:600;margin-right:8px;color:#333;display:block;margin-top:12px;";

    const providerSelect = document.createElement("select");
    providerSelect.style.cssText = "padding:6px 10px;border:1px solid #ccc;border-radius:6px;font-size:13px;margin-right:12px;";
    [
      ["claude", "Claude (all fields)"],
      ["gemini", "Gemini (all fields)"],
      ["hybrid", "Hybrid (Claude desc/details, Gemini alt)"]
    ].forEach(([value, label]) => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = label;
      if (value === CONFIG.provider) opt.selected = true;
      providerSelect.appendChild(opt);
    });
    providerSelect.addEventListener("change", () => { CONFIG.provider = providerSelect.value; });

    const altLabel = document.createElement("label");
    altLabel.textContent = "Alt Mode: ";
    altLabel.style.cssText = "font-weight:600;margin-right:8px;color:#333;display:block;margin-top:12px;";

    const altSelect = document.createElement("select");
    altSelect.style.cssText = "padding:6px 10px;border:1px solid #ccc;border-radius:6px;font-size:13px;";
    ["vision", "text"].forEach((mode) => {
      const opt = document.createElement("option");
      opt.value = mode;
      opt.textContent = mode.charAt(0).toUpperCase() + mode.slice(1);
      if (mode === CONFIG.altMode) opt.selected = true;
      altSelect.appendChild(opt);
    });
    altSelect.addEventListener("change", () => { CONFIG.altMode = altSelect.value; });

    const previewRow = document.createElement("div");
    previewRow.style.cssText = "margin-top:10px;";
    const previewCb = document.createElement("input");
    previewCb.type = "checkbox";
    previewCb.id = "ml-yext-preview";
    const previewLbl = document.createElement("label");
    previewLbl.htmlFor = "ml-yext-preview";
    previewLbl.textContent = " Preview before applying";
    previewLbl.style.cssText = "cursor:pointer;user-select:none;";
    previewRow.appendChild(previewCb);
    previewRow.appendChild(previewLbl);

    const sharedRow = document.createElement("div");
    sharedRow.style.cssText = "margin-top:6px;";
    const sharedCb = document.createElement("input");
    sharedCb.type = "checkbox";
    sharedCb.id = "ml-yext-shared-desc";
    const sharedLbl = document.createElement("label");
    sharedLbl.htmlFor = "ml-yext-shared-desc";
    sharedLbl.textContent = " Use one shared business description for all photos";
    sharedLbl.style.cssText = "cursor:pointer;user-select:none;";
    sharedRow.appendChild(sharedCb);
    sharedRow.appendChild(sharedLbl);

    const runBtn = document.createElement("button");
    runBtn.textContent = "▶ Run";
    runBtn.type = "button";
    runBtn.style.cssText = [
      "flex:1", "padding:10px 16px", "background:linear-gradient(135deg,#667eea 0%,#764ba2 100%)",
      "color:#fff", "border:none", "border-radius:8px", "cursor:pointer",
      "font-weight:600", "font-size:14px", "transition:opacity 0.15s"
    ].join(";");

    const cancelBtn = document.createElement("button");
    cancelBtn.textContent = "✕ Cancel";
    cancelBtn.type = "button";
    cancelBtn.disabled = true;
    cancelBtn.style.cssText = "padding:10px 16px;background:#e74c3c;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;font-size:14px;opacity:0.5;";
    cancelBtn.addEventListener("click", () => {
      STATE.aborted = true;
      toast("Cancelling...");
    });

    const runRow = document.createElement("div");
    runRow.style.cssText = "display:flex;gap:8px;margin-top:14px;";
    runRow.appendChild(runBtn);
    runRow.appendChild(cancelBtn);

    const clearCacheBtn = document.createElement("button");
    clearCacheBtn.textContent = "Clear Alt Cache";
    clearCacheBtn.type = "button";
    clearCacheBtn.style.cssText = "padding:6px 12px;background:#f0f0f0;border:1px solid #ccc;border-radius:6px;cursor:pointer;font-size:12px;margin-right:8px;";
    clearCacheBtn.addEventListener("click", () => {
      STATE.altState.clear();
      toast("Alt cache cleared");
    });

    const rebuildMapBtn = document.createElement("button");
    rebuildMapBtn.textContent = "Rebuild Token Map";
    rebuildMapBtn.type = "button";
    rebuildMapBtn.style.cssText = "padding:6px 12px;background:#f0f0f0;border:1px solid #ccc;border-radius:6px;cursor:pointer;font-size:12px;margin-right:8px;";
    rebuildMapBtn.addEventListener("click", () => {
      STATE.tokenMap = null;
      STATE.tokenMapBuiltAt = 0;
      getTokenMap(getEntityTitleText());
      toast("Token map rebuilt");
    });

    const debugBtn = document.createElement("button");
    debugBtn.textContent = "🔍 Debug";
    debugBtn.type = "button";
    debugBtn.style.cssText = "padding:6px 12px;background:#fffbe6;border:1px solid #e0a000;border-radius:6px;cursor:pointer;font-size:12px;";
    debugBtn.addEventListener("click", () => showDebugTrace());

    const setRunning = (running) => {
      STATE.running = running;
      runBtn.disabled = running;
      runBtn.style.opacity = running ? "0.5" : "1";
      cancelBtn.disabled = !running;
      cancelBtn.style.opacity = running ? "1" : "0.5";
    };

    runBtn.addEventListener("click", async () => {
      if (STATE.running) return;

      const mode = select.value || CONFIG.defaultMode;

      if (mode !== "sanitize") {
        const needsClaude = CONFIG.provider === "claude" || CONFIG.provider === "hybrid";
        const needsGemini = CONFIG.provider === "gemini" || CONFIG.provider === "hybrid";
        const missing = [];
        if (needsClaude && !CONFIG.getAnthropicApiKey()) missing.push("anthropicApiKey");
        if (needsGemini && !CONFIG.getGeminiApiKey()) missing.push("geminiApiKey");
        if (missing.length) {
          toast(`API key(s) not set: ${missing.join(", ")}. Use GM_setValue('<name>', '<key>') in the console.`, 8000);
          return;
        }
      }

      const scope = scopeSelect.value;
      const previewMode = previewCb.checked && mode !== "sanitize";
      const sharedDescription = sharedCb.checked && mode !== "sanitize";
      const allItems = findMediaItems();

      if (!allItems.length) {
        toast("No editable photo fields found. Open the Photo Gallery row for editing, then Run again.", 6000);
        return;
      }

      let items;
      if (scope === "single") {
        const key = singleSelect.value;
        const match = allItems.find((it) => getItemKey(it) === key);
        if (!match) {
          toast("Invalid selection. Pick an item from the dropdown.");
          return;
        }
        items = [match];
        toast(`Running on Item ${match.index + 1} only...`);
      } else if (scope === "selected") {
        items = allItems.filter((it) => SELECTED_KEYS.has(getItemKey(it)));
        if (items.length === 0) {
          toast("No photos selected. Check boxes on the photos to process.", 5000);
          return;
        }
        toast(`Running on ${items.length} selected photo(s)...`);
      } else {
        items = allItems;
      }

      removeSelectionCheckboxes();

      const entityTitle = getEntityTitleText();
      const tokenMap = getTokenMap(entityTitle);
      if (!tokenMap["[[website.url]]"]) tokenMap["[[website.url]]"] = "";

      STATE.aborted = false;
      setRunning(true);

      try {
        if (mode === "sanitize") {
          await runSanitize(items);
        } else if (mode === "overwrite") {
          await runOverwrite(items, entityTitle, tokenMap, { previewMode, sharedDescription });
        } else {
          await runSuggest(items, entityTitle, tokenMap, { previewMode, sharedDescription });
        }
      } catch (e) {
        console.error(e);
        if (e.message?.includes("cancelled")) {
          toast("Run cancelled.");
        } else {
          toast(`Error: ${e.message || e}`, 8000);
        }
      } finally {
        setRunning(false);
        STATE.aborted = false;
      }
    });

    wrap.appendChild(modeLabel);
    wrap.appendChild(select);
    wrap.appendChild(scopeLabel);
    wrap.appendChild(scopeSelect);
    wrap.appendChild(singleSelect);
    wrap.appendChild(providerLabel);
    wrap.appendChild(providerSelect);
    wrap.appendChild(altLabel);
    wrap.appendChild(altSelect);
    wrap.appendChild(previewRow);
    wrap.appendChild(sharedRow);
    wrap.appendChild(runRow);

    const buttonRow = document.createElement("div");
    buttonRow.style.cssText = "margin-top:10px;";
    buttonRow.appendChild(clearCacheBtn);
    buttonRow.appendChild(rebuildMapBtn);
    buttonRow.appendChild(debugBtn);
    wrap.appendChild(buttonRow);

    document.body.appendChild(wrap);

    const hasClaude = !!CONFIG.getAnthropicApiKey();
    const hasGemini = !!CONFIG.getGeminiApiKey();
    if (!hasClaude && !hasGemini) {
      setTimeout(() => toast("No API keys set. Use GM_setValue('anthropicApiKey', '<key>') and/or GM_setValue('geminiApiKey', '<key>') in the console.", 10000), 500);
    }
  }

  // ===========================
  // UI — DEBUG TRACE
  // ===========================

  function showDebugTrace() {
    const fromField = getNameFromNameField();
    const fromCrumb = getNameFromBreadcrumb();
    const fromPTitle = getNameFromPageTitle();
    const entityTitle = getEntityTitleText();
    const tokenMap = getTokenMap(entityTitle);
    const { name: parsedName, role: parsedRole } = extractNameRoleFromTitle(entityTitle);

    const items = findMediaItems();
    const it = items[0] || null;
    const rawDesc = it?.desc?.value || "(no desc field)";
    const rawDetails = it?.details?.value || "(no details field)";
    const rawAlt = it?.alt?.value || "(no alt field)";

    const descReverted = revertLiteralToEmbedded(rawDesc, entityTitle);
    const detailsReverted = revertLiteralToEmbedded(rawDetails, entityTitle);

    const descNorm = stripUrlTokens(descReverted);
    const isName = descIsJustEntityName(descNorm, entityTitle, rawDesc);

    const identity = it ? deriveIdentityForAlt(it, entityTitle, tokenMap) : "(no item)";

    const builtDesc = it ? buildDescriptionSuggestion(it, entityTitle) : "(no item)";
    const builtDetails = it ? buildDetailsSuggestion(it, entityTitle) : "(no item)";

    const tokenLines = Object.entries(tokenMap)
      .filter(([, v]) => v)
      .map(([k, v]) => `  ${k} = "${v}"`)
      .join("\n");

    const trace = [
      "=== ML YEXT MEDIA FULL TRACE ===",
      "",
      "-- Entity Title Sources --",
      `  fromNameField : "${fromField}"`,
      `  fromBreadcrumb: "${fromCrumb}"`,
      `  fromPageTitle : "${fromPTitle}"`,
      `  FINAL title   : "${entityTitle}"`,
      `  parsedName    : "${parsedName}"`,
      `  parsedRole    : "${parsedRole}"`,
      "",
      "-- Token Map --",
      tokenLines || "  (empty)",
      "",
      "-- Item 1 Raw --",
      `  desc    : "${rawDesc}"`,
      `  details : "${rawDetails}"`,
      `  alt     : "${rawAlt}"`,
      `  itemKey : "${it ? getItemKey(it) : "(none)"}"`,
      "",
      "-- revertLiteralToEmbedded --",
      `  desc    : "${descReverted}"`,
      `  details : "${detailsReverted}"`,
      "",
      "-- descIsJustEntityName --",
      `  descNorm : "${descNorm}"`,
      `  result   : ${isName}`,
      "",
      "-- deriveIdentityForAlt --",
      `  identity : "${identity}"`,
      "",
      "-- Built Suggestions --",
      `  description : "${builtDesc}"`,
      `  details     : "${builtDetails}"`,
      "",
      "-- Configuration --",
      `  Provider    : ${CONFIG.provider}`,
      `  Claude Key  : ${CONFIG.getAnthropicApiKey() ? "Yes" : "No (GM_setValue('anthropicApiKey', ...))"}`,
      `  Gemini Key  : ${CONFIG.getGeminiApiKey() ? "Yes" : "No (GM_setValue('geminiApiKey', ...))"}`,
      `  Claude Model: ${CONFIG.claudeModel}`,
      `  Gemini Model: ${CONFIG.geminiModel}`,
      `  Alt Mode    : ${CONFIG.altMode}`,
      `  Concurrency : ${CONFIG.itemConcurrency}`,
      "",
      "================================"
    ].join("\n");

    console.log(trace);

    let tracePanel = document.getElementById("ml-yext-trace-panel");
    if (tracePanel) tracePanel.remove();

    tracePanel = document.createElement("div");
    tracePanel.id = "ml-yext-trace-panel";
    tracePanel.style.cssText = [
      "position:fixed", "top:16px", "left:16px", "z-index:2147483647",
      "width:600px", "max-height:80vh", "overflow:auto",
      "background:#1e1e1e", "color:#d4d4d4", "font-family:monospace",
      "font-size:11px", "line-height:1.5", "padding:14px 16px",
      "border-radius:10px", "border:1px solid #444",
      "box-shadow:0 6px 22px rgba(0,0,0,0.5)", "white-space:pre"
    ].join(";");
    tracePanel.textContent = trace;

    const closeTrace = document.createElement("button");
    closeTrace.textContent = "✕ Close Trace";
    closeTrace.type = "button";
    closeTrace.style.cssText = "display:block;margin-top:10px;padding:4px 10px;cursor:pointer;background:#333;color:#fff;border:1px solid #555;border-radius:6px;font-size:11px;";
    closeTrace.addEventListener("click", () => tracePanel.remove());
    tracePanel.appendChild(closeTrace);

    document.body.appendChild(tracePanel);
  }

  // ===========================
  // BOOTSTRAP
  // ===========================

  function mediaDomReady() {
    const sel = CONFIG.selectors;
    return Boolean(
      document.querySelector(sel.description) ||
      document.querySelector(sel.details) ||
      document.querySelector(sel.alternateText) ||
      document.querySelector(sel.clickthroughUrl) ||
      document.querySelector(sel.fallbackDescription)
    );
  }

  function safeMount() {
    if (!document.getElementById("ml-yext-controls")) mountControls();
  }

  // One debounced observer handles SPA route changes AND dynamic content mounting.
  // A short retry loop covers the very-early-load case before the observer catches a mutation.
  (function bootstrap() {
    let debounceTimer = null;
    const schedule = () => {
      if (debounceTimer) return;
      debounceTimer = setTimeout(() => {
        debounceTimer = null;
        if (!document.getElementById("ml-yext-controls") && mediaDomReady()) {
          safeMount();
        }
      }, 400);
    };

    if (mediaDomReady()) safeMount();

    const obs = new MutationObserver(schedule);
    obs.observe(document.documentElement, { childList: true, subtree: true });

    const origPush = history.pushState;
    const origReplace = history.replaceState;
    history.pushState = function () {
      const r = origPush.apply(this, arguments);
      schedule();
      return r;
    };
    history.replaceState = function () {
      const r = origReplace.apply(this, arguments);
      schedule();
      return r;
    };
    window.addEventListener("popstate", schedule);

    let tries = 0;
    const maxTries = 20;
    const interval = setInterval(() => {
      tries++;
      if (document.getElementById("ml-yext-controls")) {
        clearInterval(interval);
      } else if (mediaDomReady()) {
        safeMount();
        clearInterval(interval);
      } else if (tries >= maxTries) {
        clearInterval(interval);
      }
    }, 500);
  })();
})();
