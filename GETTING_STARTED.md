# Excel Finance Agent: A Simple User Guide

A friendly walk-through to get you up and running, whether you have written
code before or not. If you get stuck anywhere, scroll to the bottom of this
guide for the "ask Claude" cheat sheet.

---

## What this thing does

You upload an Excel file. You type something like *"calculate the average of
column B"* or *"build a 5-year revenue projection at 50 lakh with 15% growth"*
in plain English. The agent figures out the right Excel formula or whole
schedule, writes it into your file, and gives you the file back.

Two ways to use it:

1. A **web app** that runs in your browser (easier for first-timers)
2. A **terminal app** if you prefer typing commands

Both do the same thing under the hood. Pick whichever feels more comfortable.

---

## What you need before starting

Check this list before you dive in. Setup takes about fifteen minutes the first
time and zero minutes after that.

- A Mac, Linux, or Windows computer
- About 200 MB of free disk space
- A working internet connection (for the AI mode; offline mode works without
  internet once installed)
- A free Mistral account for the AI mode (we will walk you through getting one)

You do NOT need:

- Any prior experience with Python, Excel macros, or programming
- A paid OpenAI / ChatGPT subscription
- An Excel licence (the file you upload can come from Google Sheets too)

---

## First-time setup

### Step 1: Open the Terminal

The Terminal is a window where you type commands instead of clicking buttons.
It is already installed on your computer.

- **Mac:** Press `Cmd` + `Space`, type "Terminal", press Enter.
- **Windows:** Press the Windows key, type "PowerShell", press Enter.
- **Linux:** You probably already know.

A black or white window opens with a prompt waiting for input. That is the
Terminal. Type the commands below into it, pressing Enter after each one.

### Step 2: Get the code

Copy and paste this into your Terminal, then press Enter:

```bash
cd ~/Desktop
git clone https://github.com/sahamate15/excel-finance-agent.git
cd excel-finance-agent
```

What just happened: you downloaded the project to a folder called
`excel-finance-agent` on your Desktop, then moved into that folder. The third
command is important. Every command from now on assumes you are inside that
folder.

If the `git clone` command fails because git is not installed, install it from
<https://git-scm.com/downloads> and try again.

### Step 3: Set up the Python sandbox

Python projects use a thing called a "virtual environment" to keep their
ingredients separate from every other Python project on your computer. Think
of it like a clean kitchen we are setting up just for this dish.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Three things just happened:

1. We made a sandboxed Python (the `.venv` folder).
2. We told the Terminal to use that sandboxed Python from now on.
3. We installed the project's ingredients (`openai`, `streamlit`, `openpyxl`,
   and so on) into the sandbox.

When the third command finishes (it can take a minute), your prompt should
have `(.venv)` at the front. That tells you the sandbox is active.

> **Windows users:** Replace `source .venv/bin/activate` with
> `.venv\Scripts\activate`.

### Step 4: Get a free Mistral API key

The agent uses Mistral, a French AI company that offers a free tier. You need
a key to use the AI features. (You can skip this step entirely if you only
want to use offline mode, but the AI mode is more capable.)

1. Go to <https://console.mistral.ai/api-keys/>
2. Sign up with your email (or Google account)
3. Click "Create new key"
4. Copy the key. It looks like a string of random letters and numbers.

Keep that key somewhere safe for the next step. Treat it like a password.

### Step 5: Save your key locally

In the Terminal, copy the example settings file to a real one:

```bash
cp .env.example .env
```

Now open `.env` in any text editor (TextEdit on Mac, Notepad on Windows, or
just `nano .env` in the Terminal) and replace `replace-me` with the key you
just copied. Save and close.

The `.env` file lives only on your computer. It is excluded from git, so it
can never accidentally end up on GitHub.

### Step 6: Confirm everything works

Run the test suite. If it shows 83 passing tests, you are good to go:

```bash
pytest tests/
```

You should see something like `83 passed in 3 seconds`. If you see failures,
scroll down to the troubleshooting section.

---

## Using the agent

### Option A: The web app (recommended for first use)

In the Terminal (with `.venv` still active), run:

```bash
streamlit run app.py
```

A browser tab opens at `http://localhost:8501`. If it does not open
automatically, copy that URL into your browser.

The page has three areas:

| Area | What it does |
|------|--------------|
| **Sidebar (left)** | Settings: mode, API key, model, dry-run toggle |
| **Run tab** | Upload a file, type instructions, generate formulas |
| **Audit Log tab** | See everything the agent has done in this session |

#### Your first formula: a quick warm-up

1. In the sidebar, leave the mode as "AI-assisted (Mistral)" and confirm the
   green checkmark next to your API key.
2. Make sure "Dry run" is checked (it is by default).
3. In the Run tab, upload the sample workbook. You can find it inside the
   project folder at `data/sample_workbook.xlsx`.
4. Pick a sheet from the dropdown. Try "Revenue".
5. In the instruction box, type: *"calculate the average of column B"*
6. Click **Generate**.

You will see a preview showing the formula the agent plans to write and
exactly which cell it will go in. Click **Confirm and write** to actually
modify the file. Click **Download updated workbook** to save it to your
Downloads folder.

Open the downloaded file in Excel, click on the new cell, and you will see
the formula. Excel calculates the result the moment you open the file.

#### Your first table: a 5-year depreciation schedule

Same flow, different instruction:

> *Create a 5-year WDV depreciation schedule for 20 lakh at 25%*

Click Generate. The dry-run preview shows you a table with all the formulas
the agent would write. Confirm to commit them. Download. Open in Excel. You
have a fully wired depreciation table where every cell is a live formula.

#### Strict mode vs AI mode

The radio button in the sidebar toggles between two operating modes.

**AI-assisted mode** sends your instruction text to Mistral. Your spreadsheet
data is never sent. Mistral receives only the words you typed. This is the
default and handles a wider range of instructions.

**Strict mode** never contacts the AI. It uses a hardcoded library of common
finance formulas plus a built-in parser for depreciation, amortization, and
projection tables. If you ask for something the offline path cannot handle,
it will tell you so loudly rather than guess.

Pick strict mode when:

- You are working on confidential deal data and your firm forbids any external
  API calls
- You want guaranteed-deterministic output
- You do not have an internet connection

Pick AI mode when:

- You want to try unusual instructions
- You need the agent to figure out which sheet, which cells, and which formula
  shape to use from a vague description

You can switch back and forth at any time. The mode change is logged.

#### Dry run: the safety net

Dry run is on by default and you should leave it on. With dry run on, the
agent shows you exactly which cells it is about to modify and what it will
write into them, then waits for you to click **Confirm and write**.

For deal models, this matters a lot. One wrong formula in a DCF and the whole
valuation is off. The dry-run preview is your last line of defence before the
file is touched.

You can turn dry run off (the toggle is in the sidebar) once you are confident
in the agent's behaviour and want a faster workflow.

#### The Audit Log tab

Click the **Audit Log** tab in the main panel to see everything the agent has
done. Every event is timestamped, hash-linked to the previous event, and
filterable by event type, file name, or source.

The **Verify integrity** button runs a tamper check. If anyone has edited the
audit log file by hand, the check catches it. The audit log records what was
done, never what was in the cells. So compliance reviewers can see "the agent
wrote =AVERAGE(B2:B100) into Revenue!F12 at 10:15 AM" but they cannot read
the underlying numbers from the audit log itself.

### Option B: The terminal (faster once you are comfortable)

If you prefer the keyboard, run:

```bash
python main.py
```

You get an interactive prompt with a Rich-styled banner. Useful commands:

| Type | What it does |
|------|--------------|
| `help` | Show example instructions |
| `file <path>` | Switch to a different workbook |
| `sheet <name>` | Switch to a different sheet |
| `preview` | Show the last few rows of the active sheet |
| `exit` or `quit` | Leave the prompt |
| Anything else | Treat as a finance instruction |

This mode does not have a built-in dry-run preview. It writes immediately. So
if you are working on a real deal model, prefer the web app.

---

## Things that might go wrong (and how to fix them)

### "ModuleNotFoundError: No module named 'openai'"

Your virtual environment is not active. Type `source .venv/bin/activate` (or
`.venv\Scripts\activate` on Windows) and try again. Your prompt should show
`(.venv)` at the front when it is active.

### "MISTRAL_API_KEY is not set"

You either skipped Step 5 or your `.env` file still says `replace-me`. Open
`.env` in a text editor, paste your real Mistral key, save, then re-run.

### "command not found: streamlit" or "command not found: python"

Same root cause as the first error: virtual environment is not active.

### Streamlit opens but the page is blank

Wait ten seconds. Streamlit takes a moment on first launch.

### The browser does not open at all

Copy the URL Streamlit prints (it looks like `http://localhost:8501`) into
your browser manually.

### Formulas show as text in the downloaded file

This is expected. The agent writes formulas; it does not calculate them.
Excel calculates the results the moment you open the file. If you are using
LibreOffice or Google Sheets, opening the file once and saving it again will
also trigger the calculation.

### A test fails when you run pytest

Three of the 83 tests need a real Mistral key to run. If only those three
fail, your key is not loading. Check `.env`. If anything else fails, scroll
down to the next section and ask Claude.

### "Permission denied" or git complains

You are probably running the commands from the wrong folder. Make sure your
prompt shows you are inside `excel-finance-agent`. If unsure, type `pwd` (Mac
or Linux) or `cd` (Windows) to see where you are.

---

## When you are truly stuck

You have Claude. Use it. Copy and paste:

1. The exact command you typed
2. The error message you got, in full
3. The output of `pwd` (so Claude knows where you are)
4. The output of `which python` (so Claude knows which Python you are using)

Ask Claude something like: *"I'm setting up this Excel Finance Agent and I
got this error. Here's what I tried and what it said. Please help."*

Most setup issues are solved in one or two messages this way.

---

## Quick reference cheat sheet

Print this section or keep it in a sticky note.

```bash
# Get into the project
cd ~/Desktop/excel-finance-agent
source .venv/bin/activate

# Run the web app
streamlit run app.py

# Run the terminal app
python main.py

# Run the tests
pytest tests/

# Query the audit log
python audit_query.py --event-type formula_written
python audit_query.py --verify $(date -u +%Y-%m-%d)
python audit_query.py --from 2026-04-01 --to 2026-04-29 --format csv

# When you are done
deactivate
```

---

## What gets written to your computer

So you know exactly what the agent touches:

| Folder or file | What it is |
|----------------|-----------|
| `data/` | Your uploaded workbooks (the agent saves a working copy here when you upload) |
| `logs/agent.log` | An operational log: a record of agent activity for debugging. Excludes raw instruction text. |
| `logs/audit/YYYY-MM-DD.jsonl` | The compliance audit log. One file per day. Tamper-evident. |
| `.env` | Your Mistral API key. Never committed to git. |
| `.venv/` | The Python sandbox. Never committed to git. |

Nothing in `logs/` or `.env` ever leaves your machine. The audit log records
what was done, not the cell values it operated on. So even if someone reviews
your audit log, they cannot reconstruct your deal data from it.

---

That is the whole guide. Welcome aboard.
