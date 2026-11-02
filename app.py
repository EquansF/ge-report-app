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
    # Formater les dates pour l'affichage
    formatted = []
    for r in reports:
        d = dict(r)
        try:
            dt = datetime.fromisoformat(r['submitted_at'].replace('Z', '+00:00'))
            d['submitted_at_fr'] = dt.strftime('%d/%m/%Y %H:%M')
        except:
            d['submitted_at_fr'] = r['submitted_at'][:16]
        formatted.append(d)
    return render_template('admin.html', reports=formatted)

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


@app.route('/admin/delete-multiple', methods=['POST'])
@login_required
def delete_multiple():
    ids = request.form.getlist('ids')
    if not ids:
        return redirect(url_for('admin'))
    conn = get_db()
    c = conn.cursor()
    placeholders = ','.join('?' * len(ids))
    c.execute(f'DELETE FROM reports WHERE id IN ({placeholders})', ids)
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

    template_path = os.path.join(os.path.dirname(__file__), 'template.xlsx')
    if not os.path.exists(template_path):
        return "template.xlsx manquant", 500

    wb = load_workbook(template_path)
    template_ws = wb.active

    for idx, r in enumerate(reports):
        data = json.loads(r['data_json'])
        if idx == 0:
            ws = template_ws
            ws.title = f"Rapport_{r['id']}"
        else:
            ws = wb.copy_worksheet(template_ws)
            ws.title = f"Rapport_{r['id']}"

        # Logo : déjà présent dans le template, on ne l'ajoute pas

        # Défusionner proprement
        for merged in list(ws.merged_cells.ranges):
            try:
                ws.unmerge_cells(str(merged))
            except Exception:
                pass

        # Date en français
        date_str = data.get('dateEssai', '')
        try:
            if date_str:
                d = datetime.strptime(date_str, '%Y-%m-%d')
                mois = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre']
                date_fr = f"{d.day} {mois[d.month-1]} {d.year}"
            else:
                date_fr = ''
        except Exception:
            date_fr = date_str

        # Remplissage
        ws['B10'] = f"COMPTE RENDU DES ESSAIS GE \nBâtiment : {data.get('batiment', '')}\nLe : {date_fr}"
        ws['F6'] = f"Rédacteur : {data.get('redacteur', '')}"
        # Même rédacteur en B34
        ws['B34'] = data.get('redacteur', '')
        ws['B19'] = f"La consommation totale lors de ces essais a été de {data.get('consoTotale', 0)} litres."
        ws['B23'] = data.get('problemes', '')

        ws['C53'] = data.get('cuve1_debut', '')
        ws['C54'] = data.get('cuve2_debut', '')

        ws['C62'] = data.get('gea_h_debut', '')
        ws['D62'] = data.get('gea_l_debut', '')
        ws['C63'] = data.get('geb_h_debut', '')
        ws['D63'] = data.get('geb_l_debut', '')
        ws['C64'] = data.get('gec_h_debut', '')
        ws['D64'] = data.get('gec_l_debut', '')
        ws['C65'] = data.get('ged_h_debut', '')
        ws['D65'] = data.get('ged_l_debut', '')

        ws['C72'] = data.get('gea_kwh_debut', '')
        ws['E72'] = data.get('gea_kvah_debut', '')
        ws['C73'] = data.get('geb_kwh_debut', '')
        ws['E73'] = data.get('geb_kvah_debut', '')
        ws['C74'] = data.get('gec_kwh_debut', '')
        ws['E74'] = data.get('gec_kvah_debut', '')
        ws['C75'] = data.get('ged_kwh_debut', '')
        ws['E75'] = data.get('ged_kvah_debut', '')

        # GE A
        ws['C84'] = data.get('gea_h_run', '')
        ws['D84'] = data.get('gea_l_run', '')
        ws['D87'] = data.get('gea_v1', '')
        ws['E87'] = data.get('gea_v2', '')
        ws['F87'] = data.get('gea_v3', '')
        ws['D88'] = data.get('gea_i1', '')
        ws['E88'] = data.get('gea_i2', '')
        ws['F88'] = data.get('gea_i3', '')
        ws['D89'] = data.get('gea_freq', '')
        ws['D90'] = data.get('gea_cos', '')
        ws['D91'] = data.get('gea_kw', '')
        ws['D92'] = data.get('gea_kva', '')
        ws['D93'] = data.get('gea_kvar', '')
        ws['D94'] = data.get('gea_kwh_run', '')
        ws['D95'] = data.get('gea_kvah_run', '')
        ws['B100'] = data.get('gea_turbo', '')
        ws['C100'] = data.get('gea_huile', '')
        ws['D100'] = data.get('gea_fioul', '')
        ws['E100'] = data.get('gea_eau', '')
        ws['F100'] = data.get('gea_ech', '')

        # GE B
        ws['C107'] = data.get('geb_h_run', '')
        ws['D107'] = data.get('geb_l_run', '')
        ws['D110'] = data.get('geb_v1', '')
        ws['E110'] = data.get('geb_v2', '')
        ws['F110'] = data.get('geb_v3', '')
        ws['D111'] = data.get('geb_i1', '')
        ws['E111'] = data.get('geb_i2', '')
        ws['F111'] = data.get('geb_i3', '')
        ws['D112'] = data.get('geb_freq', '')
        ws['D113'] = data.get('geb_cos', '')
        ws['D114'] = data.get('geb_kw', '')
        ws['D115'] = data.get('geb_kva', '')
        ws['D116'] = data.get('geb_kvar', '')
        ws['D117'] = data.get('geb_kwh_run', '')
        ws['D118'] = data.get('geb_kvah_run', '')
        ws['B123'] = data.get('geb_turbo', '')
        ws['C123'] = data.get('geb_huile', '')
        ws['D123'] = data.get('geb_fioul', '')
        ws['E123'] = data.get('geb_eau', '')
        ws['F123'] = data.get('geb_ech', '')

        # GE C
        ws['C131'] = data.get('gec_h_run', '')
        ws['D131'] = data.get('gec_l_run', '')
        ws['D134'] = data.get('gec_v1', '')
        ws['E134'] = data.get('gec_v2', '')
        ws['F134'] = data.get('gec_v3', '')
        ws['D135'] = data.get('gec_i1', '')
        ws['E135'] = data.get('gec_i2', '')
        ws['F135'] = data.get('gec_i3', '')
        ws['D136'] = data.get('gec_freq', '')
        ws['D137'] = data.get('gec_cos', '')
        ws['D138'] = data.get('gec_kw', '')
        ws['D139'] = data.get('gec_kva', '')
        ws['D140'] = data.get('gec_kvar', '')
        ws['D141'] = data.get('gec_kwh_run', '')
        ws['D142'] = data.get('gec_kvah_run', '')
        ws['B147'] = data.get('gec_turbo', '')
        ws['C147'] = data.get('gec_huile', '')
        ws['D147'] = data.get('gec_fioul', '')
        ws['E147'] = data.get('gec_eau', '')
        ws['F147'] = data.get('gec_ech', '')

        # GE D
        ws['C154'] = data.get('ged_h_run', '')
        ws['D154'] = data.get('ged_l_run', '')
        ws['D157'] = data.get('ged_v1', '')
        ws['E157'] = data.get('ged_v2', '')
        ws['F157'] = data.get('ged_v3', '')
        ws['D158'] = data.get('ged_i1', '')
        ws['E158'] = data.get('ged_i2', '')
        ws['F158'] = data.get('ged_i3', '')
        ws['D159'] = data.get('ged_freq', '')
        ws['D160'] = data.get('ged_cos', '')
        ws['D161'] = data.get('ged_kw', '')
        ws['D162'] = data.get('ged_kva', '')
        ws['D163'] = data.get('ged_kvar', '')
        ws['D164'] = data.get('ged_kwh_run', '')
        ws['D165'] = data.get('ged_kvah_run', '')
        ws['B170'] = data.get('ged_turbo', '')
        ws['C170'] = data.get('ged_huile', '')
        ws['D170'] = data.get('ged_fioul', '')
        ws['E170'] = data.get('ged_eau', '')
        ws['F170'] = data.get('ged_ech', '')

        # Fin
        ws['C179'] = data.get('cuve1_fin', '')
        ws['C180'] = data.get('cuve2_fin', '')
        ws['C187'] = data.get('gea_h_fin', '')
        ws['D187'] = data.get('gea_l_fin', '')
        ws['C188'] = data.get('geb_h_fin', '')
        ws['D188'] = data.get('geb_l_fin', '')
        ws['C189'] = data.get('gec_h_fin', '')
        ws['D189'] = data.get('gec_l_fin', '')
        ws['C190'] = data.get('ged_h_fin', '')
        ws['D190'] = data.get('ged_l_fin', '')
        ws['C196'] = data.get('gea_kwh_fin', '')
        ws['D196'] = data.get('gea_kvah_fin', '')
        ws['C197'] = data.get('geb_kwh_fin', '')
        ws['D197'] = data.get('geb_kvah_fin', '')
        ws['C198'] = data.get('gec_kwh_fin', '')
        ws['D198'] = data.get('gec_kvah_fin', '')
        ws['C199'] = data.get('ged_kwh_fin', '')
        ws['D199'] = data.get('ged_kvah_fin', '')
        ws['C204'] = data.get('gea_conso', '')
        ws['C205'] = data.get('geb_conso', '')
        ws['C206'] = data.get('gec_conso', '')
        ws['C207'] = data.get('ged_conso', '')
        ws['B210'] = f"Soit une consommation de {data.get('consoTotale', 0)} litres pendant les essais."
        ws['B213'] = data.get('commentaires', '')

        # Fusions sûres
        merges = [
            'B10:G14', 'B16:G16', 'B18:G18', 'B19:G19', 'B22:G31',
            'F5:G5', 'F6:G6', 'F7:G7', 'B210:E210',
            # Sections GE - fusions B:C sur les lignes de paramètres
            'B87:C87', 'B88:C88', 'B89:C89', 'B90:C90', 'B91:C91', 'B92:C92', 'B93:C93', 'B94:C94', 'B95:C95',
            'B110:C110', 'B111:C111', 'B112:C112', 'B113:C113', 'B114:C114', 'B115:C115', 'B116:C116', 'B117:C117', 'B118:C118',
            'B134:C134', 'B135:C135', 'B136:C136', 'B137:C137', 'B138:C138', 'B139:C139', 'B140:C140', 'B141:C141', 'B142:C142',
            'B157:C157', 'B158:C158', 'B159:C159', 'B160:C160', 'B161:C161', 'B162:C162', 'B163:C163', 'B164:C164', 'B165:C165',
            'B202:C202',
            # Fusions B:C:D demandées
            'B69:D69', 'B104:D104', 'B128:D128', 'B151:D151', 'B176:D176', 'B193:D193',
            # Nouvelles fusions
            'B212:G212', 'B213:G215',
            'B184:E184',
            'B173:G173', 'B149:G149', 'B126:G126', 'B102:G102', 'B79:G79',
            'B81:E81',
            'B46:G46', 'B50:C50', 'B59:E59',
            # Lignes 52-56 : C:D
            'C52:D52', 'C53:D53', 'C54:D54', 'C55:D55', 'C56:D56',
            # Lignes 71-76 : C:D et E:F
            'C71:D71', 'E71:F71',
            'C72:D72', 'E72:F72',
            'C73:D73', 'E73:F73',
            'C74:D74', 'E74:F74',
            'C75:D75', 'E75:F75',
            'C76:D76', 'E76:F76',
            # Zone signatures
            'B33:C33', 'D33:E33', 'F33:G33',
            'B34:C34', 'D34:E34', 'F34:G34',
            # Fusions verticales 35-38
            'B35:C38', 'D35:E38', 'F35:G38',
            'B39:C39', 'D39:E39', 'F39:G39',
            # Fusions verticales 40-43
            'B40:C43', 'D40:E43', 'F40:G43'
        ]
        for m in merges:
            try:
                ws.merge_cells(m)
            except Exception:
                pass

        # Alignements
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        bottom = Alignment(horizontal="center", vertical="bottom", wrap_text=True)

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=7):
            for cell in row:
                if cell.value is not None:
                    try:
                        cell.alignment = center
                    except Exception:
                        pass

        try:
            ws['B18'].alignment = top_left
            ws['B22'].alignment = top_left
            ws['F5'].alignment = bottom
            ws['F6'].alignment = bottom
            ws['F7'].alignment = bottom
            # Alignements top + center demandés
            top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
            for cell_ref in ['B202', 'B173', 'B149', 'B126', 'B102', 'B79']:
                ws[cell_ref].alignment = top_center
            # Alignements gauche + haut
            left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
            for cell_ref in ['B19', 'B81', 'B50', 'B59', 'B212', 'B97', 'B120', 'B144', 'B167', 'B34', 'D34', 'F34', 'B39', 'D39', 'F39']:
                try:
                    ws[cell_ref].alignment = left_top
                except:
                    pass
            # Lignes 5, 6, 7 : gauche + centre vertical (comme le modèle)
            left_center = Alignment(horizontal="left", vertical="center", wrap_text=True)
            for cell_ref in ['B5', 'B6', 'B7', 'F5', 'F6', 'F7', 'B16']:
                try:
                    ws[cell_ref].alignment = left_center
                except:
                    pass
            # Nettoyage éventuel de B6
            try:
                if ws['B6'].value in (None, ''):
                    ws['B6'].border = Border()
            except:
                pass
            # Ligne 46 : centré horizontal + vertical
            try:
                ws['B46'].alignment = center
            except:
                pass
        except Exception:
            pass

        # Bordures extérieures
        def apply_outer_border(min_row, max_row, min_col, max_col):
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    border = Border()
                    if r == min_row:
                        border.top = Side(style='thin')
                    if r == max_row:
                        border.bottom = Side(style='thin')
                    if c == min_col:
                        border.left = Side(style='thin')
                    if c == max_col:
                        border.right = Side(style='thin')
                    try:
                        cell.border = border
                    except Exception:
                        pass

        apply_outer_border(10, 14, 2, 7)  # B10:G14
        apply_outer_border(16, 20, 2, 7)  # B16:G20

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"rapports_ge_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

