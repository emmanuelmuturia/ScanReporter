#!/usr/bin/env python3
"""
Caava Group Security Scanner
ZAP scanning + Claude AI report generation → DOCX output
"""
import os
import json
import subprocess
import sys
import shutil
import re
import threading
import uuid
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
import anthropic

app = Flask(__name__, static_folder='images', static_url_path='/static/images')
app.config['UPLOAD_FOLDER'] = Path('scans')
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)
app.config['REPORTS_FOLDER'] = Path('reports')
app.config['REPORTS_FOLDER'].mkdir(exist_ok=True)

HISTORY_FILE = Path('scan_history.json')

# In-memory job store  { job_id: { status, progress, message, result } }
_jobs: dict = {}
_jobs_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Environment
# ---------------------------------------------------------------------------

def load_env():
    env_path = Path('.env')
    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, value = line.partition('=')
                key = key.strip()
                value = value.strip()
                if key and key not in os.environ:
                    os.environ[key] = value

load_env()


# ---------------------------------------------------------------------------
# Scan history helpers
# ---------------------------------------------------------------------------

def load_history():
    if HISTORY_FILE.exists():
        try:
            with open(HISTORY_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {'scans': {}}


def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)


def normalise_url(url):
    url = url.strip().rstrip('/')
    match = re.match(r'(https?://)([^/]+)(.*)', url, re.IGNORECASE)
    if match:
        url = match.group(1).lower() + match.group(2).lower() + match.group(3)
    return url


def record_scan(target_url, findings):
    history = load_history()
    key = normalise_url(target_url)
    counts = {'critical': 0, 'high': 0, 'medium': 0, 'low': 0, 'info': 0}
    for f in findings:
        sev = f.get('severity', 'info').lower()
        counts[sev] = counts.get(sev, 0) + 1
    now = datetime.utcnow().isoformat() + 'Z'
    if key not in history['scans']:
        history['scans'][key] = {
            'url': target_url, 'first_scanned': now, 'last_scanned': now,
            'scan_count': 1, 'latest': counts,
            'history': [{'scanned_at': now, 'counts': counts}]
        }
    else:
        entry = history['scans'][key]
        entry['last_scanned'] = now
        entry['scan_count'] += 1
        entry['latest'] = counts
        entry['history'].append({'scanned_at': now, 'counts': counts})
        entry['history'] = entry['history'][-10:]
    save_history(history)
    return history['scans'][key]


# ---------------------------------------------------------------------------
# ZAP Scanner
# ---------------------------------------------------------------------------

def sanitize_url_to_filename(url):
    clean = re.sub(r'https?://', '', url)
    clean = re.sub(r'[/:?&=]', '-', clean)
    clean = re.sub(r'-+', '-', clean)
    return clean.strip('-') or 'scan'


def run_zap_scan(target_url, output_dir):
    resource_name = sanitize_url_to_filename(target_url)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    scan_prefix = f"{resource_name}_{timestamp}"
    scan_json = output_dir / f"{scan_prefix}.json"
    scan_html = output_dir / f"{scan_prefix}.html"
    output_dir.chmod(0o777)

    cmd = [
        'docker', 'run', '--rm',
        '-v', f"{output_dir.absolute()}:/zap/wrk:rw",
        '--user', 'root',
        'ghcr.io/zaproxy/zaproxy:stable',
        'zap-baseline.py',
        '-t', target_url,
        '-r', scan_html.name,
        '-J', scan_json.name
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)

    if not scan_json.exists():
        stderr = result.stderr or ''
        stdout = result.stdout or ''
        raise Exception(f"ZAP scan failed: {(stderr + stdout)[:500]}")

    return {
        'json_path': scan_json,
        'html_path': scan_html,
        'resource_name': scan_prefix
    }


# ---------------------------------------------------------------------------
# ZAP JSON → structured findings
# ---------------------------------------------------------------------------

def parse_zap_findings(zap_json_path):
    with open(zap_json_path) as f:
        zap_data = json.load(f)

    alerts = zap_data.get('alerts', [])
    if not alerts and 'site' in zap_data:
        if isinstance(zap_data['site'], list):
            for site in zap_data['site']:
                alerts.extend(site.get('alerts', []))
        elif isinstance(zap_data['site'], dict):
            alerts.extend(zap_data['site'].get('alerts', []))

    severity_map = {'High': 'high', 'Medium': 'medium', 'Low': 'low', 'Informational': 'info', 'Critical': 'critical'}
    findings = []

    for alert in alerts:
        title = alert.get('alert', 'Untitled Finding')
        riskdesc = alert.get('riskdesc', 'Low (Default)')
        severity = severity_map.get(riskdesc.split(' ')[0], 'low')
        description = alert.get('desc', '').replace('<p>', '').replace('</p>', ' ').strip()
        recommendation = alert.get('solution', '').replace('<p>', '').replace('</p>', ' ').strip()

        references = []
        ref_str = alert.get('reference', '')
        if ref_str:
            refs = ref_str.replace('<p>', '').replace('</p>', '\n').split('\n')
            references = [r.strip() for r in refs if r.strip()]

        affected = []
        for instance in alert.get('instances', []):
            uri = instance.get('uri')
            param = instance.get('param')
            if uri:
                affected.append(uri + (f" [Param: {param}]" if param else ''))

        cweid = alert.get('cweid', '')
        wascid = alert.get('wascid', '')

        findings.append({
            'title': title,
            'severity': severity,
            'cvss_score': _severity_to_cvss(severity),
            'description': description,
            'recommendation': recommendation,
            'affected_components': affected or ['General Application'],
            'references': references,
            'cweid': cweid,
            'wascid': wascid,
        })

    return findings


def _severity_to_cvss(severity):
    return {'critical': '9.8', 'high': '8.1', 'medium': '6.5', 'low': '3.1', 'info': '0.0'}.get(severity, '0.0')


# ---------------------------------------------------------------------------
# Claude AI — full report generation
# ---------------------------------------------------------------------------

REPORT_SYSTEM_PROMPT = """You are a senior penetration tester and technical writer at Caava Group, \
a cybersecurity consultancy based in Nairobi, Kenya (14 Chalbi Drive, Lavington).

You write professional vulnerability assessment reports that are:
- Clear, direct, and technically accurate
- Written for both a technical and executive audience
- Structured consistently across all findings
- Compliant with CVSS 3.1 scoring and OWASP standards
- Mindful of Kenyan regulatory context (Kenya Data Protection Act 2019, KIRA where relevant)

Your reports always follow this exact structure:

1. EXECUTIVE_SUMMARY
   - overview: 2-3 paragraphs summarising the assessment, what was found, and overall risk level
   - risk_level: one of CRITICAL / HIGH / MEDIUM / LOW
   - assessment_details: dict with keys: client, target, assessment_date, assessment_type, in_scope, out_of_scope

2. FINDINGS — array of finding objects, each with:
   - id: e.g. C1, C2, H1, M1 (C=Critical, H=High, M=Medium, L=Low, I=Info)
   - title: concise vulnerability name
   - severity: critical/high/medium/low/info
   - cvss_score: numeric string e.g. "9.8"
   - cvss_vector: full CVSS 3.1 vector string
   - overview: 1-2 paragraphs explaining what it is and why it matters
   - affected_components: list of dicts with keys: component, details
   - vulnerability_details: detailed technical explanation (2-4 paragraphs)
   - proof_of_concept: step-by-step PoC description as plain text (include any request/response samples)
   - recommendations: dict with keys: immediate (list), short_term (list), long_term (list)
   - references: list of CVE/CWE/OWASP references

3. APPENDIX
   - cve_references: list of dicts with keys: id, title, cvss, description, affected_versions
   - standards: list of relevant OWASP/CWE/CVSS references
   - regulatory: list of applicable regulatory frameworks
   - disclaimer: standard disclaimer text

Respond ONLY with valid JSON matching this structure. No markdown fences, no extra text.
Add new sections or sub-sections wherever they add value beyond the raw ZAP data.
"""


def generate_report_with_claude(findings, target_url, client_name, assessment_type, scan_date):
    api_key = os.getenv('ANTHROPIC_API_KEY')
    if not api_key:
        raise Exception('ANTHROPIC_API_KEY not set in .env')

    client = anthropic.Anthropic(api_key=api_key)

    findings_summary = json.dumps(findings, indent=2)

    user_prompt = f"""Generate a complete vulnerability assessment report for the following scan.

Target URL: {target_url}
Client: {client_name}
Assessment Date: {scan_date}
Assessment Type: {assessment_type}
Assessor: Caava Group Cyber Security Team

Raw ZAP Scan Findings:
{findings_summary}

Instructions:
- Assign proper finding IDs (C1, C2... for Critical, H1... for High, M1... for Medium, L1... for Low)
- Write a compelling executive summary that references the actual findings
- For each finding, write full professional narrative — do not just echo the raw ZAP text
- Add CVSS vectors based on the nature of each vulnerability
- Add relevant CVE references where applicable
- Include OWASP Top 10 / CWE mappings
- Add Kenyan regulatory context (Kenya DPA 2019) where data privacy is at risk
- Recommendations must be tiered: Immediate (24-48h), Short Term (1-4 weeks), Long Term (1-3 months)
- The disclaimer must mention Caava Group and proper authorisation
"""

    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8000,
        system=REPORT_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}]
    )

    raw = response.content[0].text.strip()

    # Strip any accidental markdown fences
    if raw.startswith('```'):
        raw = raw.split('\n', 1)[1]
        raw = raw.rsplit('```', 1)[0].strip()

    return json.loads(raw)


# ---------------------------------------------------------------------------
# Scan job orchestration
# ---------------------------------------------------------------------------

def _run_scan_job(job_id, target_url, client_name, assessment_type, project_name):
    def update(status, progress, message):
        with _jobs_lock:
            _jobs[job_id].update(status=status, progress=progress, message=message)

    try:
        output_dir  = app.config['UPLOAD_FOLDER']
        reports_dir = app.config['REPORTS_FOLDER']

        update('running', 10, 'Starting ZAP Docker container…')
        scan_result = run_zap_scan(target_url, output_dir)

        update('running', 50, 'Parsing ZAP findings…')
        findings = parse_zap_findings(scan_result['json_path'])

        if not findings:
            update('done', 100, 'Scan complete — no findings detected.')
            with _jobs_lock:
                _jobs[job_id]['result'] = {
                    'success': True, 'target_url': target_url,
                    'findings_count': 0, 'report_file': None
                }
            return

        update('running', 65, f'Sending {len(findings)} findings to Claude…')
        scan_date = datetime.now().strftime('%B %d, %Y')
        report_data = generate_report_with_claude(
            findings, target_url, client_name, assessment_type, scan_date
        )

        # Use findings from Claude's enriched output
        enriched_findings = report_data.get('findings', findings)

        update('running', 85, 'Building DOCX report…')
        from report_builder import build_docx
        timestamp    = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_name    = re.sub(r'[^a-zA-Z0-9_-]', '_', project_name)[:40]
        report_path  = reports_dir / f"{safe_name}_{timestamp}.docx"

        build_docx(report_data, enriched_findings, target_url, client_name, report_path)

        update('running', 95, 'Recording to dashboard…')
        record_scan(target_url, enriched_findings)

        with _jobs_lock:
            _jobs[job_id].update(
                status='done', progress=100, message='Report ready',
                result={
                    'success':        True,
                    'target_url':     target_url,
                    'findings_count': len(enriched_findings),
                    'report_file':    report_path.name,
                    'scan_files': {
                        'json': str(scan_result['json_path']),
                        'html': str(scan_result['html_path']),
                    }
                }
            )

    except subprocess.TimeoutExpired:
        with _jobs_lock:
            _jobs[job_id].update(status='error',
                                 message='Scan timeout — target took too long to respond')
    except Exception as e:
        with _jobs_lock:
            _jobs[job_id].update(status='error', message=str(e))


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/scan', methods=['POST'])
def scan_endpoint():
    data            = request.get_json()
    target_url      = (data.get('url') or '').strip()
    client_name     = data.get('client_name', os.getenv('ASSESSOR_COMPANY', 'Caava Group')).strip()
    assessment_type = data.get('assessment_type', 'Web Application Security Assessment').strip()
    project_name    = data.get('project_name',
                                f"VAPT-{datetime.now().strftime('%Y%m%d-%H%M')}").strip()

    if not target_url:
        return jsonify({'error': 'URL is required'}), 400
    if not target_url.startswith(('http://', 'https://')):
        return jsonify({'error': 'URL must start with http:// or https://'}), 400

    job_id = str(uuid.uuid4())
    with _jobs_lock:
        _jobs[job_id] = {
            'status': 'running', 'progress': 0,
            'message': 'Queued…', 'result': None
        }

    threading.Thread(
        target=_run_scan_job,
        args=(job_id, target_url, client_name, assessment_type, project_name),
        daemon=True
    ).start()

    return jsonify({'job_id': job_id})


@app.route('/api/scan/status/<job_id>', methods=['GET'])
def scan_status(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify(job)


@app.route('/api/config', methods=['GET'])
def get_config():
    return jsonify({
        'has_claude_key':    bool(os.getenv('ANTHROPIC_API_KEY')),
        'docker_available':  subprocess.run(
            ['docker', '--version'], capture_output=True).returncode == 0,
        'assessor_company':  os.getenv('ASSESSOR_COMPANY', 'Caava Group'),
    })


@app.route('/api/dashboard', methods=['GET'])
def dashboard_endpoint():
    history = load_history()
    scans   = list(history.get('scans', {}).values())
    totals  = {'critical': 0, 'high': 0, 'medium': 0, 'low': 0, 'info': 0}
    for s in scans:
        for sev, n in s.get('latest', {}).items():
            totals[sev] = totals.get(sev, 0) + n
    return jsonify({
        'total_urls':  len(scans),
        'total_scans': sum(s.get('scan_count', 1) for s in scans),
        'totals':      totals,
        'urls':        sorted(scans, key=lambda x: x.get('last_scanned', ''), reverse=True)
    })


@app.route('/api/dashboard/delete/<path:url_key>', methods=['DELETE'])
def delete_dashboard_entry(url_key):
    history = load_history()
    key = normalise_url(url_key)
    if key in history.get('scans', {}):
        del history['scans'][key]
        save_history(history)
        return jsonify({'success': True})
    return jsonify({'error': 'Entry not found'}), 404


@app.route('/api/download/report/<filename>')
def download_report(filename):
    file_path = app.config['REPORTS_FOLDER'] / filename
    if file_path.exists() and file_path.suffix == '.docx':
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    return jsonify({'error': 'File not found'}), 404


@app.route('/api/download/<path:filename>')
def download_file(filename):
    file_path = app.config['UPLOAD_FOLDER'] / filename
    if file_path.exists():
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    company = os.getenv('ASSESSOR_COMPANY', 'Caava Group')
    print("\n" + "=" * 60)
    print(f"  {company} Security Scanner")
    print("=" * 60)
    print(f"\n  Starting server at http://localhost:5000")
    print(f"\n  Configuration:")
    print(f"    - Claude AI:  {'✓ Enabled' if os.getenv('ANTHROPIC_API_KEY') else '✗ Missing ANTHROPIC_API_KEY'}")
    print(f"    - Docker:     {'✓ Available' if __import__('shutil').which('docker') else '✗ Not found'}")
    print(f"\n  Reports saved to: ./reports/")
    print(f"\n  Press Ctrl+C to stop\n")
    print("=" * 60 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5000)
