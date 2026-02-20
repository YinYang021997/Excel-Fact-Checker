/* global Office, self, console, fetch, OfficeRuntime */

// Minimal type stub — avoids compile errors without full @types resolution
declare const CustomFunctions: { associate(id: string, fn: Function): void };

// ─── djb2 hash (cache key fingerprint) ───────────────────────────────────────
function djb2(str: string): string {
  let hash = 5381;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) + hash) ^ str.charCodeAt(i);
    hash = hash >>> 0; // keep unsigned 32-bit
  }
  return hash.toString(36);
}

// ─── Cache helpers (OfficeRuntime.storage, keyed by fc_<hash>) ───────────────
const CACHE_PREFIX = "fc_";

async function cacheGet(key: string): Promise<{ verdict: string; ts: number } | null> {
  try {
    const raw = await OfficeRuntime.storage.getItem(key);
    if (!raw) return null;
    return JSON.parse(raw) as { verdict: string; ts: number };
  } catch {
    return null;
  }
}

async function cacheSet(key: string, verdict: string): Promise<void> {
  try {
    await OfficeRuntime.storage.setItem(
      key,
      JSON.stringify({ verdict, ts: Date.now() })
    );
  } catch {
    // silently ignore storage errors
  }
}

// ─── Optional URL source fetch → plain text ───────────────────────────────────
async function fetchSourceText(source: string): Promise<string> {
  if (!source) return "";

  if (source.startsWith("http://") || source.startsWith("https://")) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 8000);
    try {
      const resp = await fetch(source, { signal: controller.signal });
      const html = await resp.text();
      const text = html
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, " ")
        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, " ")
        .replace(/<[^>]+>/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 6000);
      return `[Source content from ${source}]:\n${text}`;
    } catch {
      return `[Could not fetch source: ${source}]`;
    } finally {
      clearTimeout(timer);
    }
  }

  // Domain or keyword hint — pass as context only
  return `[Source context: ${source}]`;
}

// ─── Claude API ───────────────────────────────────────────────────────────────
async function callClaude(
  fact: string,
  whatItIs: string,
  sourceText: string,
  apiKey: string
): Promise<string> {
  const system = `You are a precise fact-checker. Given a claim and context, decide if the claim is accurate.
Reply with ONLY one of these three words — no punctuation, no explanation:
VERIFIED
NOT_VERIFIED
INCONCLUSIVE

VERIFIED      = claim is clearly true based on well-established knowledge or the provided source.
NOT_VERIFIED  = claim is clearly false or directly contradicted by the provided source.
INCONCLUSIVE  = cannot determine accuracy with confidence; claim is ambiguous, outdated, or needs more data.`;

  const userMsg = `Claim: "${fact}"
Context (what this is about): "${whatItIs}"
${sourceText ? "\n" + sourceText + "\n" : ""}
Reply with only VERIFIED, NOT_VERIFIED, or INCONCLUSIVE.`;

  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 32,
      system,
      messages: [{ role: "user", content: userMsg }],
    }),
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`API ${resp.status}: ${body.slice(0, 200)}`);
  }

  const data = await resp.json();
  const raw = ((data.content?.[0]?.text) ?? "").trim().toUpperCase();

  if (raw === "VERIFIED" || raw === "NOT_VERIFIED" || raw === "INCONCLUSIVE") {
    return raw;
  }
  // Claude occasionally adds punctuation — extract the keyword
  if (raw.includes("NOT_VERIFIED")) return "NOT_VERIFIED";
  if (raw.includes("VERIFIED"))     return "VERIFIED";
  return "INCONCLUSIVE";
}

// ─────────────────────────────────────────────────────────────────────────────
//  FACTCHECK — main exported custom function
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Fact-checks a claim using Claude AI.
 * Returns VERIFIED, NOT_VERIFIED, or INCONCLUSIVE.
 *
 * @customfunction
 * @param {string} fact           The claim to verify.
 * @param {string} whatItIs       Context: entity, timeframe, or metric definition.
 * @param {string} [source]       Domain (e.g. "cdc.gov") or full URL. Blank = general knowledge.
 * @param {number} [refreshHours] Cache TTL in hours. Default 24. Set to 0 to always call the API.
 * @returns {Promise<string>}     VERIFIED | NOT_VERIFIED | INCONCLUSIVE
 */
export async function FACTCHECK(
  fact: string,
  whatItIs: string,
  source?: string,
  refreshHours?: number
): Promise<string> {
  // ── Require API key ────────────────────────────────────────────────────────
  let apiKey: string | null;
  try {
    apiKey = await OfficeRuntime.storage.getItem("anthropic_api_key");
  } catch {
    return "ERROR: storage unavailable";
  }
  if (!apiKey) {
    return "ERROR: No API key — open FactCheck taskpane and save your Claude key";
  }

  // ── Cache lookup ───────────────────────────────────────────────────────────
  const ttl = refreshHours !== undefined ? refreshHours : 24;
  const cacheKey = CACHE_PREFIX + djb2(`${fact}|${whatItIs}|${source ?? ""}`);

  if (ttl > 0) {
    const hit = await cacheGet(cacheKey);
    if (hit) {
      const ageHours = (Date.now() - hit.ts) / 3_600_000;
      if (ageHours < ttl) return hit.verdict;
    }
  }

  // ── Fetch source content if a URL was supplied ─────────────────────────────
  const sourceText = source ? await fetchSourceText(source) : "";

  // ── Call Claude and cache the result ───────────────────────────────────────
  try {
    const verdict = await callClaude(fact, whatItIs, sourceText, apiKey);
    await cacheSet(cacheKey, verdict);
    return verdict;
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return "ERROR: " + msg;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  Registration — try multiple strategies in order
// ─────────────────────────────────────────────────────────────────────────────

try { (self as any)["FACTCHECK"] = FACTCHECK; } catch (_) { /* no-op */ }

function tryAssociate() {
  try {
    if (typeof CustomFunctions !== "undefined") {
      CustomFunctions.associate("FACTCHECK", FACTCHECK);
      console.log("[FactCheck] CustomFunctions.associate OK");
      return true;
    }
  } catch (e) {
    console.error("[FactCheck] CustomFunctions.associate failed:", e);
  }
  return false;
}

if (!tryAssociate()) {
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => {
      if (!tryAssociate()) {
        setTimeout(tryAssociate, 500);
      }
    });
  } else {
    let attempts = 0;
    const poll = setInterval(() => {
      if (tryAssociate() || ++attempts > 25) clearInterval(poll);
    }, 200);
  }
}
