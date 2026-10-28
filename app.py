import os
import json
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for, session, send_file, send_from_directory
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

app = Flask(__name__)
app.secret_key = 'change-this-secret-key-in-production-ge-report-2026'

# Simple admin user (change password!)
ADMIN_USERNAME = 'admin'
ADMIN_PASSWORD_HASH = generate_password_hash('admin123')  # Change this!

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

class User(UserMixin):
    def __init__(self, id):
        self.id = id

@login_manager.user_loader
def load_user(user_id):
    if user_id == '1':
        return User('1')
    return None

# Database
DB_PATH = 'ge_reports.db'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submitted_at TEXT,
            batiment TEXT,
            redacteur TEXT,
            date_essai TEXT,
            data_json TEXT,
            photo1 TEXT,
            photo2 TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        # Get JSON data
        data = request.form.get('data')
        if not data:
            return jsonify({'success': False, 'error': 'No data'}), 400
        data_dict = json.loads(data)
        
        # Handle photos
        photo1_filename = None
        photo2_filename = None
        if 'photo1' in request.files:
            photo1 = request.files['photo1']
            if photo1 and photo1.filename:
                photo1_filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_photo1_{photo1.filename}"
                photo1.save(os.path.join(app.config['UPLOAD_FOLDER'], photo1_filename))
        if 'photo2' in request.files:
            photo2 = request.files['photo2']
            if photo2 and photo2.filename:
                photo2_filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_photo2_{photo2.filename}"
                photo2.save(os.path.join(app.config['UPLOAD_FOLDER'], photo2_filename))
        
        conn = get_db()
        c = conn.cursor()
        c.execute('''
            INSERT INTO reports (submitted_at, batiment, redacteur, date_essai, data_json, photo1, photo2)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            datetime.now().isoformat(),
            data_dict.get('batiment', ''),
            data_dict.get('redacteur', ''),
            data_dict.get('dateEssai', ''),
            json.dumps(data_dict),
            photo1_filename,
            photo2_filename
        ))
        conn.commit()
        report_id = c.lastrowid
        conn.close()
        
        return jsonify({'success': True, 'id': report_id, 'message': 'Rapport enregistré avec succès !'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == ADMIN_USERNAME and check_password_hash(ADMIN_PASSWORD_HASH, password):
            user = User('1')
            login_user(user)
            return redirect(url_for('admin'))
        return render_template('login.html', error='Identifiants incorrects')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/admin')
@login_required
def admin():
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT id, submitted_at, batiment, redacteur, date_essai FROM reports ORDER BY submitted_at DESC')
    reports = c.fetchall()
    conn.close()
    return render_template('admin.html', reports=reports)

@app.route('/admin/view/<int:report_id>')
@login_required
def view_report(report_id):
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT * FROM reports WHERE id = ?', (report_id,))
    report = c.fetchone()
    conn.close()
    if not report:
        return 'Rapport non trouvé', 404
    data = json.loads(report['data_json'])
    return render_template('view.html', report=report, data=data)

@app.route('/admin/delete/<int:report_id>', methods=['POST'])
@login_required
def delete_report(report_id):
    conn = get_db()
    c = conn.cursor()
    # Optionnel : supprimer aussi les photos associées
    c.execute('SELECT photo1, photo2 FROM reports WHERE id = ?', (report_id,))
    row = c.fetchone()
    if row:
        for photo in (row['photo1'], row['photo2']):
            if photo:
                path = os.path.join(app.config['UPLOAD_FOLDER'], photo)
                if os.path.exists(path):
                    os.remove(path)
    c.execute('DELETE FROM reports WHERE id = ?', (report_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('admin'))


@app.route('/admin/export')
@login_required
def export_excel():
    ids = request.args.getlist('ids')
    conn = get_db()
    c = conn.cursor()
    if ids:
        placeholders = ','.join('?' * len(ids))
        c.execute(f'SELECT * FROM reports WHERE id IN ({placeholders}) ORDER BY submitted_at DESC', ids)
    else:
        c.execute('SELECT * FROM reports ORDER BY submitted_at DESC')
    reports = c.fetchall()
    conn.close()

    if not reports:
        return "Aucun rapport sélectionné", 400

    # Charger le modèle
    template_path = os.path.join(os.path.dirname(__file__), 'template.xlsx')
    wb = load_workbook(template_path)
    template_ws = wb.active

    # Pour plusieurs rapports, on duplique la feuille
    for idx, r in enumerate(reports):
        data = json.loads(r['data_json'])
        if idx == 0:
            ws = template_ws
            ws.title = f"Rapport_{r['id']}"
        else:
            ws = wb.copy_worksheet(template_ws)
            ws.title = f"Rapport_{r['id']}"

        # Défusionner toutes les cellules pour pouvoir écrire librement
        merged_ranges = list(ws.merged_cells.ranges)
        for merged in merged_ranges:
            ws.unmerge_cells(str(merged))

        # Remplissage des valeurs (positions basées sur le modèle fourni)
        # En-tête
        ws['E2'] = f"Rédacteur : {data.get('redacteur', '')}"
        ws['A6'] = f"COMPTE RENDU DES ESSAIS GE  Bâtiment : {data.get('batiment', '')}  Le : {data.get('dateEssai', '')}"

        # Résumé / consommation
        ws['A13'] = f"La consommation totale lors de ces essais a été de {data.get('consoTotale', 0)} litres."

        # Problèmes
        ws['A17'] = data.get('problemes', '')

        # Cuves début
        ws['B41'] = data.get('cuve1_debut', '')
        ws['B42'] = data.get('cuve2_debut', '')

        # Compteurs horaires début (GE A à D)
        base = 50
        for i, ge in enumerate(['gea', 'geb', 'gec', 'ged']):
            ws[f'B{base+i}'] = data.get(f'{ge}_h_debut', '')
            ws[f'C{base+i}'] = data.get(f'{ge}_l_debut', '')

        # Compteurs électriques début
        base = 59
        for i, ge in enumerate(['gea', 'geb', 'gec', 'ged']):
            ws[f'B{base+i}'] = data.get(f'{ge}_kwh_debut', '')
            ws[f'C{base+i}'] = data.get(f'{ge}_kvah_debut', '')

        # GE en fonctionnement - positions approximatives (à ajuster si besoin)
        # GE A
        ws['B69'] = data.get('gea_h_run', '')
        ws['C69'] = data.get('gea_l_run', '')
        ws['B71'] = data.get('gea_v1', '')
        ws['C71'] = data.get('gea_v2', '')
        ws['D71'] = data.get('gea_v3', '')
        ws['B72'] = data.get('gea_i1', '')
        ws['C72'] = data.get('gea_i2', '')
        ws['D72'] = data.get('gea_i3', '')
        ws['B73'] = data.get('gea_freq', '')
        ws['B74'] = data.get('gea_cos', '')
        ws['B75'] = data.get('gea_kw', '')
        ws['B76'] = data.get('gea_kva', '')
        ws['B77'] = data.get('gea_kvar', '')
        ws['B78'] = data.get('gea_kwh_run', '')
        ws['B79'] = data.get('gea_kvah_run', '')
        ws['A82'] = data.get('gea_turbo', '')
        ws['B82'] = data.get('gea_huile', '')
        ws['C82'] = data.get('gea_fioul', '')
        ws['D82'] = data.get('gea_eau', '')
        ws['E82'] = data.get('gea_ech', '')

        # GE B (décalé)
        offset = 20
        ws[f'B{69+offset}'] = data.get('geb_h_run', '')
        ws[f'C{69+offset}'] = data.get('geb_l_run', '')
        # ... (on peut étendre de la même façon pour C et D)

        # Fin - cuves
        # À adapter selon les lignes exactes du modèle
        ws['B120'] = data.get('cuve1_fin', '')
        ws['B121'] = data.get('cuve2_fin', '')

        # Consommations
        ws['B130'] = data.get('gea_conso', '')
        ws['B131'] = data.get('geb_conso', '')
        ws['B132'] = data.get('gec_conso', '')
        ws['B133'] = data.get('ged_conso', '')
        ws['A135'] = f"Soit une consommation de {data.get('consoTotale', 0)} litres pendant les essais."

        # Commentaires
        ws['A138'] = data.get('commentaires', '')

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"rapports_ge_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/uploads/<filename>')
@login_required
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
