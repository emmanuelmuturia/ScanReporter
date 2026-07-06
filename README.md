# Caava Group Security Scanner

Automated web application security scanner that runs a ZAP scan, feeds the findings to Claude AI, and generates a fully branded **Caava Group DOCX vulnerability report** — in one click.

## Flow

```
Target URL → ZAP Docker Scan → Parse Findings → Claude AI → Branded DOCX Report
```

## Setup

### 1. Clone and create virtual environment

```bash
git clone <your-repo>
cd scrappy-scanner
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env`:

```
ANTHROPIC_API_KEY=sk-ant-...          # Required — get from console.anthropic.com
ASSESSOR_NAME=Your Name
ASSESSOR_COMPANY=Caava Group
ASSESSOR_ADDRESS=14 Chalbi Drive, Lavington\nNairobi
```

### 3. Add your logo (optional)

Place `caava_logo.png` in the `images/` folder.
The report builder will embed it on the cover page.
If no logo file is found, it falls back to a text header.

### 4. Run

```bash
python app.py
```

Open http://localhost:5000

## What Claude generates

Claude receives all ZAP findings plus your client/assessment details and produces:

- **Executive Summary** — overview, risk level, assessment details table, in/out of scope
- **Identified Vulnerabilities** — CVSS-scored table with severity chart
- **Per-Finding Sections** — overview, affected components, technical details, PoC, tiered recommendations (Immediate / Short Term / Long Term)
- **Appendix** — CVE references, OWASP/CWE mappings, Kenyan regulatory context (DPA 2019, KIRA), disclaimer

Claude is free to add new sections where the findings warrant it.

## Output

Reports are saved to `./reports/` as `.docx` files and are available for download immediately from the UI.

## Prerequisites

- Docker (for ZAP)
- Python 3.9+
- Anthropic API key
