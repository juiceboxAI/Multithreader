# AGENTS.md — [MULTI]THREADER (Codex Operating Manual)

You are the lead coder + design engineer for **[MULTI]THREADER**, a local-first browser tool for people who live in tabs: task-maxers, multi-tool workflows, rapid notes + sheet-like data, fast context switching, minimal friction.

This repo is intentionally small and direct: a single-page app with no backend, and **local persistence** (localStorage + optional cache-file handle where supported). The codebase centers around `app.js` and a single in-memory `state` tree. :contentReference[oaicite:0]{index=0}

---

## 0) Prime Directive

Deliver a tool that feels like:

- **Instant** (no laggy UI rituals)
- **Predictable** (undo never surprises, paste never vandalizes your sheet)
- **Local & private** (no network calls, no tracking, no “sign in” nonsense)
- **Low-drama** (clean UX, sane shortcuts, stable data model)

If a feature threatens predictability, we either:

1. redesign it until it’s predictable, or
2. ship it behind a toggle, or
3. don’t ship it.

---

## 1) Non-Negotiables (Guardrails)

### Data safety

- No destructive actions without confirmation (clear, delete thread, replace sheet on import).
- State mutations must be followed by `scheduleSave()` (and must not thrash).
- Never silently discard user data.

### Local-first integrity

- Do not add network calls.
- Keep storage format backward compatible when possible:
  - prefer `version` fields + migration paths over breaking changes.

### UI predictability

- “Undo” must always undo the last user-intentful action in the _same surface_ (Scheme vs Data).
- Paste into a cell must behave like users expect from spreadsheets (and must not scatter text unless it’s clearly a grid paste).

---

## 2) Current Architecture (What exists today)

### Single state tree (persisted)

- `state = { activeId, threads[], ui{} }`
- saved to localStorage under `STORAGE_KEY = "multithreader.v1"` via `scheduleSave()`.
- optional cache-file writes via File System Access API with handle stored in IndexedDB (`multithreader-cache`). :contentReference[oaicite:1]{index=1}

### Threads

Each thread contains:

- `journal` (Scheme text)
- `sheet` (Data grid: `cells[r][c] = { value, bg, align }`)
- per-thread UI preferences (fonts, sizes, colors)

### Undo model (today)

- Journal undo stacks: `journalUndoStacks` keyed by thread id
- Sheet undo stacks: `sheetUndoStacks` keyed by thread id
- limit: `UNDO_LIMIT = 10` (global constant)

### Data grid editing (today)

- Each cell is a text `<input>` with a custom paste handler.
- Formula bar exists (`formulaInput`), and formulas are evaluated on display:
  - supported functions: `SUM/AVG/MIN/MAX/COUNT`
  - cell refs like `A1`, ranges like `A1:B3`

### Context menus (today)

- Tabs have a context menu (add/delete thread).
- Sheet has a context menu for **row/col headers only** (insert/remove row/col).  
  There is **no per-cell context menu yet**. :contentReference[oaicite:2]{index=2}

### Break timer (today)

- Break timer counts down, shows modal + tone on completion.
- There is a “pulse” reminder every 30 minutes using a fixed interval (not escalating, not activity-aware). :contentReference[oaicite:3]{index=3}

---

## 3) Known Issues (must address)

### Issue A — Memory: “4 tabs ≈ 75MB”

Treat this as **unknown until profiled**.

- Browser baselines vary wildly; 75MB for a JS-heavy SPA may be normal.
- But we must rule out “we’re accidentally retaining old DOM / closures / timers.”

Acceptance:

- Add a lightweight internal “diagnostics” panel or dev-only logging:
  - thread count
  - total cells
  - active timers (breakTickTimer, breakReminderTimer)
  - undo stack sizes
- Run Chrome Performance/Memory profile:
  - verify no runaway listeners, no duplicate intervals, no detached DOM trees.

Do **not** prematurely micro-opt unless profiling shows a leak.

### Issue B — Pasting long text into a cell “splats everywhere”

Root cause (likely): cell paste handler treats any newline/tab/comma as a grid and applies across multiple cells. This is correct for TSV/CSV pastes, but wrong for “a big blob of text”.

Fix strategy:

- Paste classification rules for **cell paste**:
  1. If clipboard contains `\t` → treat as grid (TSV) ✅
  2. Else if clipboard looks like CSV (commas + multiple rows) → treat as grid ✅
  3. Else treat as **single cell**, even if it contains newlines ✅
- Optional power-user behavior:
  - If user holds `Alt` (or another modifier) while pasting: force grid mode.

Acceptance:

- Multi-line text paste lands entirely in one cell.
- TSV paste still fans out across cells.

### Issue C — Undo feels buggy / crosses surfaces / clears formats

We need to make undo semantics boring and reliable.

Probable contributors:

- Multiple pushes per single intent (blur + input + formula bar commit).
- Recalc/display overwriting visible content (formula results vs raw) can make undo _look_ like it broke even when state is correct.
- Format changes (shade/align) are stored in the same undo snapshots as value changes; users may perceive “undo cleared formatting” unexpectedly.

Fix strategy:

- Split undo stacks by “surface” (already separate: journal vs sheet) **and make it impossible to call the wrong undo**:
  - Ctrl+Z should undo the currently focused surface only.
- Define “transactional undo” for sheet:
  - One user action = one undo entry.
  - For selection operations (shade/align), push exactly once.
  - For cell edits: push on _edit begin_ or _edit commit_, not both.
- Store and display raw value vs rendered display more explicitly:
  - in-sheet inputs should show raw when focused, rendered when blurred (already), but ensure blur commits don’t push twice.

Acceptance:

- Undo in Data never affects Scheme.
- Undo restores both value and formatting predictably.
- Undo does not randomly “clear a field” unless that was the last action.

---

## 4) Additions Requested (build next)

### A) Cell-level right click menu (Data + Scheme field)

Implement a **unified context menu** pattern:

- Right click on Data cell / Scheme field shows:
  - Copy / Cut / Paste
  - Text color (future: per-cell text color; today we only have bg + alignment)
  - Cell background color (existing concept via `bg`)
  - Alignment: left/center/right (existing)
  - Wrap / Truncate toggle (new)
  - Add formula (Data tile → set value to `=` and focus formula bar)

Behavior:

- Pressing Enter after choosing a color applies it.
- For selection: actions apply to selection range.

Notes:

- Text color does not exist in the cell model yet. If we add it:
  - extend cell object: `{ value, bg, align, fg?, wrap? }`
  - migration: default `fg=""` and `wrap="wrap"` or boolean.

### B) Arrow key navigation in the sheet

Expected behavior:

- Arrow keys move selection cell-to-cell when not editing.
- When editing:
  - arrows move cursor inside input (standard text input behavior).
- Enter:
  - commit edit and move down (optional) or stay (choose one, document it).

### C) Data Tile improvements

You called out “Data Tile” specifically; in the current app, the grid is the tile.
Implement:

- wrap/truncate per-cell or per-selection
- “right click → add formula”
- (optional) a tiny “fx” affordance for formula cells

---

## 5) Changes Requested (break reminder + tab indicator)

### Break reminder escalation (new behavior)

Goal: noticeable but not annoying.

Rules:

- Gentle reminder is activity-aware:
  - only ticks while the page is active (visible)
  - resets when a break is started (and does not re-start until the break timer ends)
- Escalation:
  - 30 min: soft green pulse
  - 60 min: soft yellow pulse
  - 90 min: soft red pulse
  - 120 min: steady glowing red until a break starts
- When break timer ends:
  - play a non-startling alert tone
  - show modal/popup notification
  - snooze in 10-minute increments
- Reset must reset _all_ timers: break countdown + escalation reminder.

Implementation notes:

- Current reminder uses `setInterval(..., 30*60*1000)` + `pulseBreakButton()`; replace with a 1-minute ticker and a computed “elapsed focus time”.
- Track:
  - `focusStartAt` (timestamp when page became active)
  - `accumulatedFocusMs` (carryover)
  - `reminderStage` (0/30/60/90/120)
- Use `document.visibilitychange` to pause accumulation when hidden.

### Active tab indicator (visual)

Requirement: active thread’s tab shows indicator ON; inactive darkened.

Today: `.tab.active` exists. Add:

- a small “status dot” element inside each tab
- CSS:
  - active = bright
  - inactive = dim
- Keep it purely visual (no logic side effects).

---

## 6) Explore (deeper context per thread)

“Tabs in the data sheet / pages in scheme window up to 3”

Treat as an experiment branch:

- Add optional per-thread “pages”:
  - `thread.journalPages: [string,string,string]` with active index
  - `thread.sheetPages: [sheet,sheet,sheet]` with active index
- Keep v1 compatibility by defaulting to single page.
- UX: small page pills “1 2 3” within each panel.

Do not ship this until:

- paste + undo are stable (otherwise we multiply chaos).

---

## 7) Coding Standards (Codex rules)

### Style

- Prefer small pure functions.
- Avoid “clever” abstractions.
- Name state transitions clearly: `startBreakTimer`, `resetBreakTimer`, `applyGridToSheet`, etc.
- Any new state fields must be normalized in `normalizeState()` / `normalizeSheet()`.

### DOM + events

- Never attach duplicate global listeners.
- Any `setInterval` must have a corresponding stop/cleanup path.

### Performance

- No re-rendering the entire table unless size changes.
- Keep operations on selections O(n) in selection size, not O(total cells) when possible.

---

## 8) Definition of Done (DoD)

A change is “done” when:

- It’s implemented
- It doesn’t break local storage restore
- It doesn’t regress paste / undo / selection
- It has at least a quick manual test script (steps listed in PR/commit notes)
- It feels consistent with the rest of the UI

---

## 9) Immediate Work Order (recommended execution order)

1. Fix cell paste classification (stop “splat”)
2. Stabilize undo semantics (surface-isolated + transactional)
3. Add arrow navigation
4. Implement cell context menu (copy/paste, bg, align, wrap/truncate, add formula)
5. Break reminder escalation + reset logic + visibility awareness
6. Active tab indicator polish
7. Explore: 3 pages per panel (optional)

---
