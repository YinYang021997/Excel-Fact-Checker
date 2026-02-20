# Excel Fact Checker

An Excel add-in that adds a `=FC.FACTCHECK()` custom function powered by Claude AI. Type a claim into any cell and get back **VERIFIED**, **NOT_VERIFIED**, or **INCONCLUSIVE** — with optional URL source fetching and 24-hour result caching.

```
=FC.FACTCHECK("New York is the capital of USA", "US geography")
→ NOT_VERIFIED
```

---

## Prerequisites

Before starting, make sure you have these installed:

| Tool | Download |
|------|----------|
| **Node.js** (v18 or later) | https://nodejs.org |
| **Git** | https://git-scm.com |
| **Microsoft Excel** (Microsoft 365 desktop, Windows or Mac) | Microsoft 365 subscription |
| **Claude API key** | https://console.anthropic.com → API Keys |

---

## One-Time Setup

Do these steps **once** on a new machine.

### 1. Clone the repo

```bash
git clone https://github.com/YinYang021997/Excel-Fact-Checker.git
cd Excel-Fact-Checker
```

### 2. Install dependencies

```bash
npm install
```

### 3. Trust the local HTTPS certificate

The add-in runs over HTTPS on localhost. Run this once to install a trusted dev certificate:

```bash
npx office-addin-dev-certs install
```

If prompted by Windows/Mac for admin permission, allow it.

### 4. Register the add-in with Excel

```bash
npm run start
```

This registers the add-in in Excel's developer registry so it appears automatically every time Excel opens. You only need to run this **once** — you can close Excel again afterwards.

> **Note:** If Excel opens and shows an error on first run, close it. The registration still happened. Proceed to the next step.

### 5. Build the project

```bash
npm run build:dev
```

---

## Every-Session Usage

Each time you want to use the add-in, follow these steps:

### Step 1 — Start the dev server

Open a terminal in the project folder and run:

```bash
npm run dev-server
```

Keep this terminal open the entire time you use Excel. The add-in loads its files from this local server.

### Step 2 — Open Excel

Open Excel normally. The **FactCheck** group should appear in the **Home** ribbon automatically.

### Step 3 — Save your API key (first time, or after clearing storage)

1. Click **FactCheck Settings** in the Home ribbon
2. Paste your Claude API key (`sk-ant-api03-…`) into the API Key field
3. Click **Save Key** — the badge should change to "Saved (…XXXX)"

Your key is stored locally in Excel's secure storage and persists across sessions.

### Step 4 — Use the formula

```excel
=FC.FACTCHECK(fact, whatItIs, [source], [refreshHours])
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `fact` | string | Yes | The claim to verify |
| `whatItIs` | string | Yes | Context — what entity, timeframe, or metric this is about |
| `source` | string | No | A full URL (fetched live) or domain hint (e.g. `"cdc.gov"`). Leave blank to use Claude's training knowledge |
| `refreshHours` | number | No | Cache TTL in hours. Default: `24`. Set to `0` to always call the API |

**Examples:**

```excel
=FC.FACTCHECK("The Eiffel Tower is in Paris", "landmark location")
→ VERIFIED

=FC.FACTCHECK("New York is the capital of USA", "US geography")
→ NOT_VERIFIED

=FC.FACTCHECK(A2, B2)
→ checks the claim in A2 with context from B2

=FC.FACTCHECK(A2, B2, "https://en.wikipedia.org/wiki/New_York_City")
→ fetches the Wikipedia page and uses it as the source

=FC.FACTCHECK(A2, B2, "", 0)
→ bypasses cache, always calls the API fresh
```

**Returns:** `VERIFIED` | `NOT_VERIFIED` | `INCONCLUSIVE` | `ERROR: <message>`

### Step 5 — Stop when done

Press `Ctrl+C` in the terminal running `npm run dev-server`, or just close the terminal.

---

## Troubleshooting

### Cells show `#VALUE!` after opening Excel

The shared runtime needs a few seconds to initialize. Wait 3–4 seconds after Excel opens, then press:

```
Ctrl + Alt + F9
```

This forces Excel to recalculate all custom function cells.

### "No API key" error in cell

Open the FactCheck taskpane (Home ribbon → **FactCheck Settings**) and save your Claude API key.

### ADD-IN ERROR dialog appears

Click **Close** on the dialog (not Restart). The formula should still work. If it persists:
1. Close Excel completely
2. Delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` (Windows) — these are safe cache files
3. Reopen Excel and start the dev server again

### Port 3000 already in use

Another process is holding port 3000. Find and kill it:

```bash
# Windows
netstat -ano | findstr :3000
taskkill /PID <the-pid> /F

# Mac/Linux
lsof -i :3000
kill -9 <the-pid>
```

Then run `npm run dev-server` again.

---

## Project Structure

```
factcheck-addin/
├── src/
│   ├── functions/
│   │   ├── functions.ts        # FACTCHECK custom function (Claude API logic)
│   │   ├── functions.json      # Function metadata (parameters, types)
│   │   └── functions.html      # Minimal runtime page for the functions iframe
│   └── taskpane/
│       ├── taskpane.html       # Taskpane UI template
│       └── taskpane.ts         # API key management + cache controls
├── assets/                     # Add-in icons (16×16, 32×32, 80×80)
├── manifest.xml                # Office add-in manifest
├── webpack.config.js           # Build configuration
├── package.json
└── tsconfig.json
```

---

## Useful Commands

| Command | What it does |
|---------|-------------|
| `npm run dev-server` | Start the local HTTPS dev server (required every session) |
| `npm run build:dev` | Rebuild bundles in development mode |
| `npm run build` | Rebuild bundles in production mode (minified) |
| `npm run start` | Register add-in with Excel (one-time setup only) |
| `npm run validate` | Validate the manifest XML |

---

## How It Works

1. Excel evaluates `=FC.FACTCHECK(...)` and calls the registered JavaScript function
2. The function checks `OfficeRuntime.storage` for a cached result (keyed by a hash of the inputs)
3. If no valid cache entry exists, it optionally fetches the source URL, then calls the Claude API
4. Claude returns `VERIFIED`, `NOT_VERIFIED`, or `INCONCLUSIVE`
5. The result is cached and written to the cell

API calls use the `claude-sonnet-4-6` model with a 32-token response limit (just the verdict word), keeping costs minimal.
