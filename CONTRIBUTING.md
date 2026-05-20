# Contributing to outlook-desktop-mcp

## Branching Strategy

```
feature/your-change → PR → preview → PR → main → auto-publish to PyPI
```

- **`main`** — stable, auto-publishes to PyPI on every push. Never push directly.
- **`preview`** — integration testing branch. All PRs land here first.
- **Feature branches** — your working branches, created from `preview`.

## How to Contribute

1. **Fork** the repo to your GitHub account
2. **Clone** your fork locally
3. **Create a branch** from `preview`:
   ```bash
   git checkout preview
   git pull origin preview
   git checkout -b feature/my-change
   ```
4. **Make your changes**, commit, push to your fork
5. **Open a PR** into `preview` (not `main`)
6. Once reviewed and merged to `preview`, it will be tested there
7. Periodically, `preview` is merged into `main` which triggers a PyPI release

## Development Setup

Unit tests (no Outlook required) run on Windows, macOS, and Linux.
Integration testing requires the matching desktop Outlook.

### Windows (with Outlook Desktop Classic for integration tests)

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
python .venv\Scripts\pywin32_postinstall.py -install
```

### macOS / Linux (unit tests only)

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

Add `".[cli]"` instead of `".[dev]"` if you want the bundled `mcp` CLI tools at runtime.

## Testing

### Unit tests (no Outlook required — gates PyPI publish)

```bash
pip install -e ".[dev]"
pytest                         # runs tests/unit by default
# or, on Windows:
outlook-desktop-mcp.cmd test-unit
```

CI runs the unit suite on ubuntu/macOS/windows with Python 3.10–3.13 on every
push and PR. `publish.yml` will not push to PyPI unless the unit suite passes
**and** `pyproject.toml`'s version is greater than the current PyPI release.

### Integration tests (require real Outlook)

With Classic Outlook running on Windows:

```bash
set RUN_OUTLOOK_INTEGRATION=1
pytest tests/contacts_mcp_test.py tests/contacts_com_test.py -v
python tests\phase3_mcp_test.py   # legacy script-style validators
outlook-desktop-mcp.cmd test
```

The legacy `phase1_com_test.py`/`calendar_com_test.py`/`extras_com_test.py`
files are not collected by pytest (see `tests/conftest.py`); run them directly
with `python tests/<name>.py`.

Contact tools cache successful results for 7 days (max 256 entries, LRU).
Failed `resolve_recipient` and empty macOS lists are not cached. Restart the
MCP server to refresh after editing contacts in Outlook.

## Adding New Tools

1. Define the COM function that does the work (receives `outlook, namespace` as first args)
2. Add an `@mcp.tool()` async handler in `server.py` that calls `bridge.call(your_function, ...)`
3. Write a detailed docstring — this is what LLMs see during tool discovery
4. Add test coverage
5. Update the `instructions` string if adding a new capability category
