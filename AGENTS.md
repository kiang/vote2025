# Repository Guidelines

## Project Structure & Module Organization
- `excel_to_cunli_converter.py`: Parses CEC Excel files in `raw/`, matches to VILLCODE, and writes `docs/referendum_cunli_data.json` with verified totals.
- `compare_referendums.py`: Compares 2025 village totals with 2021 (case 17) and outputs `referendum_comparison_2025_vs_2021.csv`.
- `raw/`: Place source `.xlsx` files (one per county). Example: `raw/縣表3-臺北市-全國性公民投票.xlsx`.
- `docs/`: Built artifacts published via GitHub Pages (see `.github/workflows/static.yml`).
- `manual_villcode_mapping.json`: Optional overrides and multi-village splits; `unmatched_for_mapping.json` is generated to guide updates.

## Build, Test, and Development Commands
- Install deps (Python 3.10+): `python3 -m venv .venv && . .venv/bin/activate && pip install pandas`.
- Generate JSON from Excel: `python3 excel_to_cunli_converter.py` (reads `raw/*.xlsx`, writes `docs/referendum_cunli_data.json`).
- Compare 2025 vs 2021: `python3 compare_referendums.py` (requires 2021 data at `/home/kiang/public_html/referendums2021/docs/data.json`). Adjust the path if needed.
- Preview Pages locally: `python3 -m http.server -d docs 8080`.

## Coding Style & Naming Conventions
- Python: 4-space indentation, snake_case for files, functions, and variables; descriptive names for data fields.
- I/O: Use UTF-8, avoid absolute paths unless necessary; prefer `Path`/relative paths where possible.
- Optional tools: `black` and `ruff` for formatting/linting. Keep functions small and pure; document assumptions.

## Testing Guidelines
- No test suite yet. If adding tests, use `pytest` under `tests/` with files named `test_*.py`.
- Focus tests on: row parsing, VILLCODE matching (including multi-village splits), and total verification.
- Aim for basic coverage on critical paths; include sample minimal `.xlsx` fixtures if feasible.

## Commit & Pull Request Guidelines
- Commits: Imperative, concise subjects (<= 50 chars). Example: "Fix vote counting in shared stations". Add a body for context and rationale when changing logic.
- PRs: Include a clear description, linked issue (if any), before/after notes for totals, and sample command outputs. Attach screenshots or CSV snippets when helpful.
- CI: Pages deploy publishes `docs/`. Do not commit large raw datasets; place them in `raw/` locally.

## Security & Configuration Tips
- Large files: Avoid committing raw Excel files; they are not needed for Pages.
- Paths: `compare_referendums.py` uses an absolute 2021 data path—update it or parameterize before running on other hosts.
