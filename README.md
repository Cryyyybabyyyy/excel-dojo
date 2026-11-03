# Excel Dojo V2.0 — Learn · Practice · Master

An offline-first static learning website with six Excel learning modules and **60 hands-on formula challenges**.

## Run on Windows

1. Extract `Excel_Dojo_V2_Website.zip` into a folder.
2. Open `index.html` in Edge or Chrome. Keep `index.html`, `curriculum.js`, `engine.js`, `app.js`, and `styles.css` together.
3. If browser privacy settings block local storage, launch a local server from the extracted folder:

```powershell
py -m http.server 8000
```

Visit http://localhost:8000

You do not need Node.js or Python for normal local use; the Python option is for a development server.

## Publish online

Upload all five web files to the root of a GitHub repository, then select **Settings → Pages → Deploy from a branch → main / root**. The site is static; no build step or server-side database is required.

OneDrive can store the ZIP or HTML files, but a OneDrive sharing link is not equivalent to static site hosting.

## Curriculum

1. Spreadsheet Fundamentals — arithmetic, ranges, references, invoice boss.
2. Essential Excel Functions — SUM, AVERAGE, MIN, MAX, COUNT, COUNTA and rounding.
3. Percentages & Finance — discount, growth, VAT, margins, conditional pricing.
4. Logic & Decision Making — IF, AND, OR, NOT, IFS, boundaries.
5. Conditional Calculations — COUNTIF(S), SUMIF(S), AVERAGEIF(S), criteria.
6. Errors & Formula Debugging — IFERROR, IFNA, ISNUMBER, ISTEXT, ISBLANK, ISERROR.

Each challenge includes a lesson, worked example, interactive dataset exploration, an editable spreadsheet, specific feedback, three hints, an answer reveal, six formula validation datasets, and a concept mastery question. Multi-formula boss challenges grade all target cells together.

## Important limitations

- The application implements a **safe, deliberately limited Excel-like formula interpreter**. It supports the functions used by its learning curriculum, not the entire Excel calculation language. It does not execute VBA, macros, scripts or arbitrary JavaScript.
- The formula grader assesses results over a small set of generated data variations and enforces selected required techniques; this is not a mathematical proof of equivalence. Some alternative valid Excel solutions may not be recognised.
- The lesson 'Explore' input panel currently edits numeric source cells. Text scenarios are shown as worksheet examples.
- Mastery checks are conceptual three-choice questions, not a full adaptive assessment engine.
- Progress is saved with browser `localStorage`, not synced between devices. Use **Achievements → Export progress** and **Import progress** to back it up. The V1 progress file is not compatible with V2.
- A text reference to an Excel error type is educational; Excel's full coercion and error semantics are not duplicated exactly.

## Technical design

- `curriculum.js` — module descriptions, 60 problems, input datasets, reference formulas, quizzes, hints.
- `engine.js` — tokenizer, parser, safe formula interpreter, grading, constrained reference checks.
- `app.js` — routing, interactive lessons, practice workspace, XP/mastery and persistent state.
- `styles.css` — responsive Forest Green & Sage design.
- `index.html` — application shell.

Primary palette: forest `#3B6939`, sage `#F7FAF0`, white `#FFFFFF`, charcoal `#181D16`.


## V2.1 Concept Lessons Upgrade
All 60 concept pages now use their own beginner-friendly lesson explanation, sample data, step-by-step worked formula, challenge-specific pitfall, and reflection prompt. The expandable sidebar, practice workspace, grader, mastery checks and saved progress are unchanged.
