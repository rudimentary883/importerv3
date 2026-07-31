"""
Tabbycat API Importer - Pure Python, no pandas
Works on any Python version including 3.14
Optimized for Render Free Tier (512MB RAM)
"""

import os
import io
import csv
import json
import time
import requests

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'tabbycat-importer-key-2024')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


def clean_string(val):
    if val is None:
        return ''
    s = str(val).strip()
    return s if s != 'None' else ''


def parse_bool(val):
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    return str(val).strip().upper() in ('TRUE', '1', 'YES', 'Y', 'T')


def parse_float_or_none(val):
    if val is None or str(val).strip() == '':
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def read_csv_file(file):
    """Read CSV file into list of dicts using stdlib only"""
    text = file.read().decode('utf-8')
    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


def read_excel_file(file):
    """Read Excel file using openpyxl"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ImportError("openpyxl is required for Excel files")
    
    wb = load_workbook(file)
    ws = wb.active
    
    headers = [cell.value for cell in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_dict = {}
        for i, header in enumerate(headers):
            if header:
                row_dict[header] = row[i] if i < len(row) else None
        rows.append(row_dict)
    return rows


def read_uploaded_file(file):
    ext = file.filename.rsplit('.', 1)[1].lower()
    if ext == 'csv':
        return read_csv_file(file)
    return read_excel_file(file)


# =============================================================================
# TABBYCAT API CLIENT
# =============================================================================

class TabbycatAPI:
    def __init__(self, base_url, token, tournament_slug):
        self.base_url = base_url.rstrip('/')
        self.token = token
        self.slug = tournament_slug
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Token {token}',
            'Content-Type': 'application/json'
        })
        self.created_institutions = {}
        self.stats = {'success': 0, 'failed': 0, 'errors': []}

    def _url(self, path):
        return f"{self.base_url}/api/tournaments/{self.slug}{path}"

    def _post(self, path, data, retries=3):
        url = self._url(path)
        for attempt in range(retries):
            try:
                time.sleep(0.3)
                resp = self.session.post(url, json=data, timeout=30)
                if resp.status_code == 201:
                    self.stats['success'] += 1
                    return resp.json()
                elif resp.status_code == 429:
                    time.sleep(2 ** attempt)
                    continue
                else:
                    error = f"HTTP {resp.status_code}: {resp.text[:200]}"
                    self.stats['errors'].append(error)
                    self.stats['failed'] += 1
                    return None
            except requests.exceptions.RequestException as e:
                if attempt == retries - 1:
                    error = f"Request failed: {str(e)}"
                    self.stats['errors'].append(error)
                    self.stats['failed'] += 1
                    return None
                time.sleep(1)
        return None

    def create_institution(self, name, code):
        if code in self.created_institutions:
            return self.created_institutions[code]
        data = {"name": name, "code": code}
        result = self._post('/institutions/', data)
        if result and 'id' in result:
            self.created_institutions[code] = result['id']
            return result['id']
        return None

    def create_team(self, institution_id, reference, short_reference,
                    use_institution_prefix=True, emoji='', speakers=None, code_name=''):
        data = {
            "institution": institution_id,
            "reference": reference,
            "short_reference": short_reference or reference,
            "use_institution_prefix": use_institution_prefix,
        }
        if emoji:
            data["emoji"] = emoji
        if code_name:
            data["code_name"] = code_name
        if speakers:
            data["speakers"] = speakers
        return self._post('/teams/', data)

    def create_adjudicator(self, name, institution_id=None, email='',
                          gender='', base_score=None, independent=False,
                          adj_core=False, notes=''):
        data = {"name": name}
        if institution_id:
            data["institution"] = institution_id
        if email:
            data["email"] = email
        if gender:
            data["gender"] = gender
        if base_score is not None:
            data["base_score"] = base_score
        data["independent"] = independent
        data["adj_core"] = adj_core
        if notes:
            data["notes"] = notes
        return self._post('/adjudicators/', data)

    def test_connection(self):
        try:
            url = f"{self.base_url}/api/"
            resp = self.session.get(url, timeout=10)
            return resp.status_code == 200
        except Exception:
            return False


# =============================================================================
# CSV PROCESSORS (pure Python)
# =============================================================================

def process_institutions(rows):
    results = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        name = clean_string(row.get('name', ''))
        code = clean_string(row.get('code', ''))
        if not name and not code:
            continue
        if not name:
            errors.append(f"Row {idx}: Missing institution name")
            continue
        if not code:
            errors.append(f"Row {idx}: Missing code for '{name}'")
            continue
        results.append({'name': name, 'code': code})
    return results, errors


def process_adjudicators(rows):
    results = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        name = clean_string(row.get('name', ''))
        if not name:
            continue
        gender = clean_string(row.get('gender', ''))
        gender_norm = ''
        if gender.upper() in ['M', 'MALE']:
            gender_norm = 'M'
        elif gender.upper() in ['F', 'FEMALE']:
            gender_norm = 'F'
        elif gender.upper() in ['O', 'OTHER']:
            gender_norm = 'O'
        results.append({
            'name': name,
            'institution': clean_string(row.get('institution', '')),
            'email': clean_string(row.get('email', '')),
            'gender': gender_norm,
            'base_score': parse_float_or_none(row.get('base_score')),
            'independent': parse_bool(row.get('independent')),
            'adj_core': parse_bool(row.get('adj_core')),
            'notes': clean_string(row.get('notes', ''))
        })
    return results, errors


def process_teams(rows):
    results = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        institution = clean_string(row.get('institution', ''))
        if not institution:
            continue
        ref = clean_string(row.get('reference', ''))
        if not ref:
            errors.append(f"Row {idx}: Missing reference for institution '{institution}'")
            continue
        short_ref = clean_string(row.get('short_reference', ref))
        results.append({
            'institution': institution,
            'reference': ref,
            'short_reference': short_ref or ref,
            'code_name': clean_string(row.get('code_name', '')),
            'use_institution_prefix': parse_bool(row.get('use_institution_prefix', True)),
            'emoji': clean_string(row.get('emoji', '')),
            'team_name_human': clean_string(row.get('team_name (human)', ''))
        })
    return results, errors


def process_speakers(rows):
    results = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        name = clean_string(row.get('name', ''))
        if not name:
            continue
        gender = clean_string(row.get('gender', ''))
        gender_norm = ''
        if gender.upper() in ['M', 'MALE']:
            gender_norm = 'M'
        elif gender.upper() in ['F', 'FEMALE']:
            gender_norm = 'F'
        elif gender.upper() in ['O', 'OTHER']:
            gender_norm = 'O'
        results.append({
            'name': name,
            'gender': gender_norm,
            'email': clean_string(row.get('email', '')),
            'phone': clean_string(row.get('phone', '')),
            'anonymous': parse_bool(row.get('anonymous')),
            'team': clean_string(row.get('team', '')),
            'categories': clean_string(row.get('categories', '')),
            'initials_match': clean_string(row.get('initials_match', ''))
        })
    return results, errors


# =============================================================================
# CSV GENERATORS
# =============================================================================

def generate_institutions_csv(institutions):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['name', 'code'])
    for inst in institutions:
        writer.writerow([inst['name'], inst['code']])
    return output.getvalue()


def generate_adjudicators_csv(adjudicators):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['institution', 'name', 'email', 'gender', 'base_score', 'independent', 'adj_core', 'notes'])
    for adj in adjudicators:
        writer.writerow([
            adj['institution'], adj['name'], adj['email'], adj['gender'],
            adj['base_score'] if adj['base_score'] is not None else '',
            'TRUE' if adj['independent'] else 'FALSE',
            'TRUE' if adj['adj_core'] else 'FALSE',
            adj['notes']
        ])
    return output.getvalue()


def generate_teams_csv(teams):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['institution', 'reference', 'short_reference', 'use_institution_prefix', 'emoji'])
    for team in teams:
        writer.writerow([
            team['institution'], team['reference'], team['short_reference'],
            'TRUE' if team['use_institution_prefix'] else 'FALSE',
            team['emoji']
        ])
    return output.getvalue()


def generate_speakers_csv(speakers):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['team', 'name', 'email', 'phone', 'gender', 'anonymous'])
    for spk in speakers:
        writer.writerow([
            spk['team'], spk['name'], spk['email'], spk['phone'],
            spk['gender'], 'TRUE' if spk['anonymous'] else 'FALSE'
        ])
    return output.getvalue()


# =============================================================================
# FLASK ROUTES
# =============================================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/test-connection', methods=['POST'])
def test_connection():
    data = request.get_json()
    api = TabbycatAPI(
        data.get('base_url', ''),
        data.get('token', ''),
        data.get('slug', '')
    )
    ok = api.test_connection()
    return jsonify({'ok': ok})


@app.route('/upload', methods=['POST'])
def upload():
    mode = request.form.get('mode', 'csv')

    try:
        if 'institutions' not in request.files:
            flash('Institutions file is required', 'error')
            return redirect(url_for('index'))

        inst_file = request.files['institutions']
        adj_file = request.files.get('adjudicators')
        teams_file = request.files.get('teams')
        speakers_file = request.files.get('speakers')

        if inst_file.filename == '':
            flash('Institutions file is required', 'error')
            return redirect(url_for('index'))

        inst_rows = read_uploaded_file(inst_file)
        institutions, inst_errors = process_institutions(inst_rows)
        institution_codes = {i['code'] for i in institutions}

        adjudicators = []
        adj_errors = []
        teams = []
        team_errors = []
        speakers = []
        speaker_errors = []

        if adj_file and adj_file.filename:
            adj_rows = read_uploaded_file(adj_file)
            adjudicators, adj_errors = process_adjudicators(adj_rows)

        if teams_file and teams_file.filename:
            team_rows = read_uploaded_file(teams_file)
            teams, team_errors = process_teams(team_rows)
            for t in teams:
                if t['institution'] not in institution_codes:
                    team_errors.append(f"Team '{t['institution']} {t['reference']}': institution '{t['institution']}' not found")

        if speakers_file and speakers_file.filename:
            speaker_rows = read_uploaded_file(speakers_file)
            speakers, speaker_errors = process_speakers(speaker_rows)

        api_results = None
        if mode == 'api':
            base_url = request.form.get('api_url', '').strip()
            token = request.form.get('api_token', '').strip()
            slug = request.form.get('tournament_slug', '').strip()

            if not all([base_url, token, slug]):
                flash('API URL, Token, and Tournament Slug are required for API mode', 'error')
                return redirect(url_for('index'))

            api = TabbycatAPI(base_url, token, slug)

            if not api.test_connection():
                flash('Could not connect to Tabbycat API. Check URL, token, and that API is enabled.', 'error')
                return redirect(url_for('index'))

            for inst in institutions:
                api.create_institution(inst['name'], inst['code'])

            if speakers:
                speakers_by_team = {}
                for spk in speakers:
                    t = spk['team']
                    if t not in speakers_by_team:
                        speakers_by_team[t] = []
                    spk_data = {"name": spk['name']}
                    if spk['email']:
                        spk_data["email"] = spk['email']
                    if spk['gender']:
                        spk_data["gender"] = spk['gender']
                    spk_data["anonymous"] = spk['anonymous']
                    speakers_by_team[t].append(spk_data)

                for team in teams:
                    inst_id = api.created_institutions.get(team['institution'])
                    team_name = team['team_name_human'] or f"{team['institution']} {team['reference']}"
                    team_speakers = speakers_by_team.get(team_name, [])
                    api.create_team(
                        institution_id=inst_id,
                        reference=team['reference'],
                        short_reference=team['short_reference'],
                        use_institution_prefix=team['use_institution_prefix'],
                        emoji=team['emoji'],
                        code_name=team['code_name'],
                        speakers=team_speakers if team_speakers else None
                    )
            else:
                for team in teams:
                    inst_id = api.created_institutions.get(team['institution'])
                    api.create_team(
                        institution_id=inst_id,
                        reference=team['reference'],
                        short_reference=team['short_reference'],
                        use_institution_prefix=team['use_institution_prefix'],
                        emoji=team['emoji'],
                        code_name=team['code_name']
                    )

            for adj in adjudicators:
                inst_id = api.created_institutions.get(adj['institution']) if adj['institution'] else None
                api.create_adjudicator(
                    name=adj['name'],
                    institution_id=inst_id,
                    email=adj['email'],
                    gender=adj['gender'],
                    base_score=adj['base_score'],
                    independent=adj['independent'],
                    adj_core=adj['adj_core'],
                    notes=adj['notes']
                )

            api_results = api.stats

        inst_csv = generate_institutions_csv(institutions)
        adj_csv = generate_adjudicators_csv(adjudicators)
        teams_csv = generate_teams_csv(teams)
        speakers_csv = generate_speakers_csv(speakers)

        from flask import session
        session['institutions_csv'] = inst_csv
        session['adjudicators_csv'] = adj_csv
        session['teams_csv'] = teams_csv
        session['speakers_csv'] = speakers_csv

        return render_template('results.html',
                               institutions=institutions,
                               inst_errors=inst_errors,
                               adjudicators=adjudicators,
                               adj_errors=adj_errors,
                               teams=teams,
                               team_errors=team_errors,
                               speakers=speakers,
                               speaker_errors=speaker_errors,
                               inst_count=len(institutions),
                               adj_count=len(adjudicators),
                               team_count=len(teams),
                               speaker_count=len(speakers),
                               mode=mode,
                               api_results=api_results)

    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download/<file_type>')
def download(file_type):
    from flask import session
    files = {
        'institutions': ('institutions.csv', session.get('institutions_csv', '')),
        'adjudicators': ('adjudicators.csv', session.get('adjudicators_csv', '')),
        'teams': ('teams.csv', session.get('teams_csv', '')),
        'speakers': ('speakers.csv', session.get('speakers_csv', ''))
    }
    if file_type not in files:
        flash('Invalid file type', 'error')
        return redirect(url_for('index'))

    filename, content = files[file_type]
    buffer = io.BytesIO(content.encode('utf-8'))
    return send_file(buffer, mimetype='text/csv', as_attachment=True, download_name=filename)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
