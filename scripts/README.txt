Development checks
==================
Run from the repository root, with Python 3:
  python scripts/build_site.py --bundles
  python tests/run_tests.py
  python tests/test_learning.py

The build refreshes the full-text catalog, file allowlist, two legacy-named course ZIP files and SHA-256 manifest. Rebuild after changing content. The separate media lab and historical development reports are not included in course ZIPs.

Optional browser checks require Node.js, Playwright and Chromium. The checked browser path is /usr/bin/google-chrome; adapt it to the local installation if needed:
  node tests/browser_regression.cjs
Set AQAR_EVIDENCE to a writable directory for screenshots and browser-results.json. The suite runs a loopback-only HTTP server and closes it at the end. It does not change a hosted site.

Workbook generator:
  python نماذج/مولد-النموذج-المالي.py
It uses only the Python standard library, and writes the workbook beside the script (نماذج/نموذج-دراسة-الجدوى-العقارية.xlsx) regardless of the directory it is invoked from. After any workbook regeneration, rebuild bundles and rerun the tests. The independent cached-value recalculation in this review used LibreOffice; the Python evaluator tests only its supported formula subset. Spreadsheet software must recalculate formulas when opening/editing the workbook.

These checks are regressions, not certification of every legal statement or market assumption. Read the dated audit report for coverage and limits.
