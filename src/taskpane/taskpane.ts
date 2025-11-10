/* global Office, Word */

/**
 * Option B: party-aware highlighter (fast & robust)
 * - Detect parties (N), render radio buttons
 * - Next: highlight selected party across body:
 *     name (+ possessive), role aliases (“Company”, “the Company”, …),
 *     and pronouns (it/its/itself OR they/their/themselves)
 * - Reset: clear only our add-in highlights
 * - Case-insensitive; tolerant search for longer terms
 * - Batching + pronoun throttling for speed
 */

type Party = { name: string; role?: string };

const TAG = "AI_PARTY_HL";
const STORAGE_KEY_PARTIES = "ai.parties";
const STORAGE_KEY_SELECTED = "ai.selectedParty";
const HIGHLIGHT_COLOR = "Turquoise";

// Performance knobs
const READ_TEXT_LIMIT = 12000;        // read more of the doc header for detection
const BATCH_SIZE = 60;                // wrap N hits then sync
const PRONOUN_HIT_CAP = 250;          // skip terms with too many matches (e.g., “it”)
const STOPWORDS = /^(services?|deliverables?|agreement|parties?)$/i;

// DOM
let elParties: HTMLElement;
let elPartyManual: HTMLInputElement;
let elBtnAddParty: HTMLButtonElement;
let elBtnRedetect: HTMLButtonElement;
let elBtnNext: HTMLButtonElement;
let elBtnReset: HTMLButtonElement;
let elBusy: HTMLElement;
let elBusyText: HTMLElement;
let elStatus: HTMLElement;

Office.onReady(() => {
  elParties     = document.getElementById("parties")!;
  elPartyManual = document.getElementById("partyManual") as HTMLInputElement;
  elBtnAddParty = document.getElementById("btnAddParty") as HTMLButtonElement;
  elBtnRedetect = document.getElementById("btnRedetect") as HTMLButtonElement;
  elBtnNext     = document.getElementById("btnNext") as HTMLButtonElement;
  elBtnReset    = document.getElementById("btnReset") as HTMLButtonElement;
  elBusy        = document.getElementById("busy")!;
  elBusyText    = document.getElementById("busyText")!;
  elStatus      = document.getElementById("status")!;

  elBtnAddParty.addEventListener("click", onAddParty);
  elBtnRedetect.addEventListener("click", async () => detectAndRenderParties(true));
  elBtnNext.addEventListener("click", onNext);
  elBtnReset.addEventListener("click", onReset);

  initPane().catch(console.error);
});

/* ---------------- UI helpers ---------------- */

function setBusy(on: boolean, msg?: string) {
  if (on) {
    if (msg) elBusyText.textContent = msg;
    elBusy.removeAttribute("hidden");
  } else {
    elBusy.setAttribute("hidden", "true");
  }
}

function setStatus(text: string) {
  elStatus.textContent = text;
}

function renderParties(parties: Party[], selectedName?: string) {
  elParties.innerHTML = "";
  if (parties.length === 0) {
    elParties.innerHTML = `<div class="hint">No parties detected yet. Add one or click <b>Redetect</b>.</div>`;
    return;
  }
  parties.forEach((p, idx) => {
    const id = `party-${idx}`;
    const div = document.createElement("div");
    div.className = "party-item";
    div.innerHTML = `
      <input type="radio" name="party" id="${id}" value="${escapeHtml(p.name)}" ${selectedName === p.name ? "checked" : ""} />
      <label for="${id}">${escapeHtml(p.name)}${p.role ? ` <span class="hint">(${escapeHtml(p.role)})</span>` : ""}</label>
    `;
    elParties.appendChild(div);
  });
}

function clearPartySelectionUI() {
  const checked = elParties.querySelector<HTMLInputElement>('input[name="party"]:checked');
  if (checked) checked.checked = false;
}

function getSelectedPartyName(): string | undefined {
  const input = elParties.querySelector<HTMLInputElement>('input[name="party"]:checked');
  return input?.value;
}

function escapeHtml(s: string) {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string));
}

/* --------------- lifecycle --------------- */

async function initPane() {
  const [storedParties, storedSelected] = await getStoredState();
  if (storedParties && storedParties.length) {
    renderParties(storedParties, storedSelected);
  } else {
    await detectAndRenderParties(false);
  }
  setStatus("Ready.");
}

async function getStoredState(): Promise<[Party[] | null, string | undefined]> {
  return new Promise((resolve) => {
    const parties  = Office.context.document.settings.get(STORAGE_KEY_PARTIES) as Party[] | null;
    const selected = Office.context.document.settings.get(STORAGE_KEY_SELECTED) as string | undefined;
    resolve([parties, selected]);
  });
}

async function saveParties(parties: Party[]) {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.settings.set(STORAGE_KEY_PARTIES, parties);
    Office.context.document.settings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

async function saveSelected(name: string | undefined) {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.settings.set(STORAGE_KEY_SELECTED, name || null);
    Office.context.document.settings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

/* --------- party detection / mgmt --------- */

async function detectAndRenderParties(_force: boolean) {
  setBusy(true, "Detecting parties…");
  try {
    const text = await readDocumentHeadText(READ_TEXT_LIMIT);
    let parties = detectParties(text)
      .filter((p) => !STOPWORDS.test(p.name.trim()));     // basic false-positive filter
    parties = dedupeParties(parties);                     // case-insensitive de-dupe

    await saveParties(parties);
    const selected = parties[0]?.name;
    await saveSelected(selected);
    renderParties(parties, selected);

    setStatus(parties.length ? `Detected ${parties.length} party(ies).` : "No parties detected.");
  } catch (e) {
    console.error(e);
    setStatus("Detection failed. Add manually or try again.");
  } finally {
    setBusy(false);
  }
}

function onAddParty() {
  const name = (elPartyManual.value || "").trim();
  if (!name) return;

  getStoredState().then(async ([parties]) => {
    const list = dedupeParties([...(parties || []), { name }]);
    await saveParties(list);
    await saveSelected(name);
    renderParties(list, name);
    setStatus(`Selected party: ${name}`);
    elPartyManual.value = "";
  });
}

function equalName(a: string, b: string) {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

function dedupeParties(arr: Party[]) {
  const seen = new Set<string>();
  const out: Party[] = [];
  for (const p of arr) {
    const key = p.name.trim().toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      out.push(p);
    }
  }
  return out;
}

/* ------------ Word interactions ------------ */

async function readDocumentHeadText(limitChars = READ_TEXT_LIMIT): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const txt = body.text || "";
    return txt.slice(0, Math.max(0, limitChars));
  });
}

async function onNext() {
  const selected = getSelectedPartyName();
  if (!selected) {
    setStatus("Pick a party first.");
    return;
  }

  setBusy(true, "Clearing previous highlights…");
  await clearPartyHighlights();

  setBusy(true, "Building terms…");
  const [parties] = await getStoredState();
  const party = (parties || []).find((p) => equalName(p.name, selected)) || { name: selected };

  const aliases = buildAliases(party);
  const pronouns = pronounsFor(party.name);
  const terms = buildTerms(party.name, aliases, pronouns).sort((a, b) => b.length - a.length);

  await saveSelected(party.name);

  let totalHits = 0;
  for (let i = 0; i < terms.length; i++) {
    const term = terms[i];
    setBusy(true, `Highlighting “${term}” (${i + 1}/${terms.length})…`);
    const hits = await searchAndHighlight(term, !isPossessive(term));
    totalHits += hits;
  }

  setBusy(false);
  setStatus(`Highlighted ${totalHits} occurrence(s) for ${party.name}.`);
}

async function onReset() {
  setBusy(true, "Removing highlights…");
  await clearPartyHighlights();
  await saveSelected(undefined);
  clearPartySelectionUI();
  setBusy(false);
  setStatus("All add-in highlights removed. Party selection cleared.");
}

// remove ONLY our tagged content controls
async function clearPartyHighlights() {
  return Word.run(async (context) => {
    const ccs = context.document.body.contentControls.getByTag(TAG);
    ccs.load("items");
    await context.sync();

    ccs.items.forEach((cc) => {
      cc.font.highlightColor = null;
      cc.delete(true);
    });

    await context.sync();
  });
}

/**
 * Fast & tolerant highlight:
 *  - Ignores punctuation/space only for longer terms
 *  - Skips very noisy short terms (e.g., “it”) if they exceed PRONOUN_HIT_CAP
 *  - Processes hits in batches to minimize syncs
 *  - If a batch errors (overlap etc.), falls back to per-item for that batch
 */
async function searchAndHighlight(term: string, wholeWord: boolean): Promise<number> {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Build search options (cast to any to avoid TS typing issues for ignore flags)
    const opts: any = { matchCase: false, matchWholeWord: wholeWord };
    if (term.length > 3) {
      opts.ignorePunct = true;
      opts.ignoreSpace = true;
    }
    let results = body.search(term, opts);
    results.load("items");
    await context.sync();

    // Throttle extremely noisy terms (pronouns)
    if (results.items.length > PRONOUN_HIT_CAP && term.length <= 3) {
      // “it”/“its”/“they”… would be too many → skip to keep the UI responsive
      return 0;
    }

    // Process in batches
    let applied = 0;
    for (let start = results.items.length - 1; start >= 0; ) {
      const end = Math.max(-1, start - BATCH_SIZE);
      const slice: Word.Range[] = [];
      for (let i = start; i > end; i--) slice.push(results.items[i]);

      // Try whole slice in one go
      try {
        for (const r of slice) {
          const cc = r.insertContentControl();
          cc.tag = TAG;
          cc.font.highlightColor = HIGHLIGHT_COLOR;
        }
        await context.sync();
        applied += slice.length;
        start = end;
        continue;
      } catch {
        // Fallback: per item for this batch
        for (const r of slice) {
          try {
            const cc = r.insertContentControl();
            cc.tag = TAG;
            cc.font.highlightColor = HIGHLIGHT_COLOR;
            await context.sync();
            applied++;
          } catch {
            // Overlap or odd range—skip it
          }
        }
        start = end;
      }
    }

    return applied;
  });
}

/* --------------- term building --------------- */

function buildTerms(name: string, aliases: string[], pronouns: string[]) {
  const unique = new Set<string>();
  add(unique, name);
  add(unique, makePossessive(name));
  for (const a of aliases) {
    add(unique, a);
    add(unique, makePossessive(a));
  }
  for (const p of pronouns) add(unique, p);
  return Array.from(unique);
}

function add(set: Set<string>, s?: string) {
  if (s && s.trim()) set.add(s.trim());
}

function makePossessive(s: string) {
  return !s ? s : s.endsWith("s") || s.endsWith("S") ? `${s}'` : `${s}'s`;
}

function buildAliases(p: Party): string[] {
  const out: string[] = [];
  const base = ["Company","Contractor","Client","Customer","Licensor","Licensee","Provider","Vendor","Reseller","Partner"];
  const addForms = (r: string) => out.push(r, `the ${r}`, r.toUpperCase(), `the ${r.toUpperCase()}`);
  if (p.role) addForms(p.role); else base.forEach(addForms);
  return dedupe(out);
}

function pronounsFor(name: string) {
  const pluralHint = /(holdings|group|partners)\b/i.test(name) || /[^']s\b/i.test(name);
  return pluralHint ? ["they", "their", "themselves"] : ["it", "its", "itself"];
}

function isPossessive(term: string) {
  return /'\s*s?$/.test(term);
}

function dedupe<T>(arr: T[]) {
  return Array.from(new Set(arr));
}

/* ---------------- party detection ---------------- */

function detectParties(text: string): Party[] {
  const t = normalizeSpaces(text);
  const candidates: Party[] = [];

  // “between X and Y …”
  const m = t.match(/\b(?:by and )?between\s+(.{2,140}?)\s+and\s+(.{2,140}?)[\.;,\n]/i);
  if (m) {
    const a = scrubName(m[1]); if (a) candidates.push(extractRole(a));
    const b = scrubName(m[2]); if (b) candidates.push(extractRole(b));
  }

  // Fallback: scan early lines for org-like names
  if (candidates.length < 2) {
    const orgs = findOrgNames(t);
    for (const o of orgs) {
      if (!candidates.find((p) => equalName(p.name, o))) candidates.push(extractRole(o));
      if (candidates.length >= 2) break;
    }
  }

  return candidates.slice(0, 5);
}

function extractRole(fragment: string): Party {
  const roleMatch = fragment.match(/\(\s*(?:the\s+)?["“](Company|Contractor|Client|Customer|Licensor|Licensee|Provider|Vendor|Reseller|Partner)["”]\s*\)/i);
  const role = roleMatch ? titleCase(roleMatch[1]) : undefined;
  const name = fragment.replace(/\(\s*(?:the\s+)?["“](?:Company|Contractor|Client|Customer|Licensor|Licensee|Provider|Vendor|Reseller|Partner)["”]\s*\)/ig, "").trim();
  return { name: stripQuotes(name), role };
}

function findOrgNames(t: string): string[] {
  const tokens = /(Inc\.?|Incorporated|Corp\.?|Corporation|LLC|L\.?L\.?C\.?|Ltd\.?|Limited|Holdings|Group|Partners|Company)/i;
  const lines = t.split(/\n+/).slice(0, 40);
  const out: string[] = [];
  for (const ln of lines) {
    if (tokens.test(ln)) {
      const m = ln.match(/([A-Z][A-Za-z0-9&.,'’\- ]{2,80}?(?:Inc\.?|Incorporated|Corp\.?|Corporation|LLC|L\.?L\.?C\.?|Ltd\.?|Limited|Holdings|Group|Partners|Company))/);
      if (m) out.push(scrubName(m[1]));
    }
  }
  return dedupe(out);
}

function scrubName(s: string) {
  let x = s.replace(/\b(?:by and )?between\b/i, "");
  x = x.replace(/^[,;:\-\s]+|[,;:\-\s]+$/g, "");
  x = x.replace(/\s{2,}/g, " ");
  return stripQuotes(x);
}

function stripQuotes(s: string) {
  return s.replace(/^[“"']|[”"']$/g, "").trim();
}

function titleCase(s: string) {
  return s.replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
}

function normalizeSpaces(s: string) {
  return s.replace(/\r\n?/g, "\n").replace(/[ \t]+/g, " ").trim();
}
