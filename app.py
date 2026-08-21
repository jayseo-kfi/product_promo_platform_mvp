import os
import json
import sqlite3
from datetime import datetime
from functools import wraps
from io import BytesIO
from werkzeug.middleware.proxy_fix import ProxyFix

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file,
    send_from_directory,
    abort
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Oracle Cloud에서는 DATA_DIR 환경변수로 영구 데이터 저장 위치를 지정합니다.
# 로컬 컴퓨터에서는 프로젝트 폴더 아래 data 디렉터리를 사용합니다.

DATA_DIR = os.environ.get(
    'DATA_DIR',
    os.path.join(BASE_DIR, 'data')
)

DB_PATH = os.path.join(DATA_DIR, 'platform.db')
UPLOAD_DIR = os.path.join(DATA_DIR, 'uploads')

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.wsgi_app = ProxyFix(
    app.wsgi_app,
    x_for=1,
    x_proto=1,
    x_host=1
)

app.secret_key = os.environ.get('SECRET_KEY')
if not app.secret_key:
    raise RuntimeError('SECRET_KEY 환경변수가 설정되지 않았습니다.')

app.config.update(
    MAX_CONTENT_LENGTH=100 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    # HTTPS 적용 후 Oracle 서버 환경변수에서 1로 설정합니다.
    SESSION_COOKIE_SECURE=os.environ.get('SESSION_COOKIE_SECURE', '0') == '1',
)

CATEGORIES = ['건축','구조','기계설비','전기','통신','소방','토목','조경','보안','에너지','서비스','기타']
COMPANY_CERTS = ['여성기업','사회적기업','사회적협동조합','장애인기업','중증장애인생산품 생산시설','장애인표준사업장','중소기업','창업기업','기타']
POLICY_PRODUCTS = ['중소기업제품','기술개발제품','여성기업제품','사회적기업 생산품','사회적협동조합 생산품','장애인기업 생산품','중증장애인 생산품','장애인표준사업장 생산품','녹색제품','창업기업제품','상생협력제품','혁신제품 또는 혁신시제품','시범구매제품','해당 없음','기타']
CERT_TYPES = ['KS 인증','KC 인증','녹색기술 인증','고효율에너지기자재 인증','환경표지 인증','신제품 인증','신기술 인증','성능인증','우수조달물품','품질경영시스템 인증','안전 관련 인증','특허 또는 실용신안','정부권장정책 이행제품','시험성적서','기타']

ALLOWED_EXTENSIONS = {
    # PDF
    'pdf',

    # 한글
    'hwp', 'hwpx',

    # Word
    'doc', 'docx',

    # Excel
    'xls', 'xlsx', 'xlsm', 'csv',

    # PowerPoint
    'ppt', 'pptx',

    # 텍스트
    'txt',

    # 이미지
    'jpg', 'jpeg', 'png', 'gif', 'bmp', 'tif', 'tiff',

    # 압축파일
    'zip'
}


def allowed_file(filename):
    return (
        '.' in filename
        and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def save_file(file):
    if not file or not file.filename:
        return None

    if not allowed_file(file.filename):
        abort(400, description='허용되지 않는 파일 형식입니다.')

    safe_name = secure_filename(file.filename)
    name = datetime.now().strftime('%Y%m%d%H%M%S%f') + '_' + safe_name
    path = os.path.join(UPLOAD_DIR, name)
    file.save(path)
    return name



def db():
    conn = sqlite3.connect(
        DB_PATH,
        timeout=30
    )

    conn.row_factory = sqlite3.Row

    conn.execute('PRAGMA foreign_keys = ON')
    conn.execute('PRAGMA busy_timeout = 30000')
    conn.execute('PRAGMA journal_mode = WAL')

    return conn


def format_company_certs(certs_json, other=''):
    try:
        certs = json.loads(certs_json or '[]')
    except (TypeError, json.JSONDecodeError):
        certs = []
    labels = [str(x) for x in certs if x and x != '기타']
    if '기타' in certs and other:
        labels.append(f'기타({other})')
    return ', '.join(labels) if labels else '해당 없음'


@app.template_filter('company_categories')
def company_categories_filter(value):
    if isinstance(value, dict):
        return format_company_certs(value.get('company_certs'), value.get('company_cert_other'))
    return format_company_certs(value)

def init_db():
    conn = db()
    try:
        conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            role TEXT NOT NULL DEFAULT 'vendor',
            business_no TEXT UNIQUE,
            company_name TEXT UNIQUE NOT NULL,
            company_certs TEXT,
            company_cert_other TEXT,
            contact_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            privacy_consent INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            last_login TEXT,
            active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            product_name TEXT NOT NULL,
            category TEXT NOT NULL,
            category_other TEXT,
            manufacturer TEXT NOT NULL,
            model_name TEXT,
            country TEXT NOT NULL,
            policy_products TEXT,
            policy_other TEXT,
            size_spec TEXT,
            performance TEXT,
            lifespan TEXT,
            applicable_facilities TEXT,
            compatibility TEXT,
            delivery_period TEXT,
            spec_file TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS certifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            cert_type TEXT NOT NULL,
            cert_name TEXT,
            agency TEXT,
            cert_no TEXT,
            issue_date TEXT,
            valid_from TEXT,
            valid_to TEXT,
            no_expiry INTEGER NOT NULL DEFAULT 0,
            file_path TEXT,
            note TEXT,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS deliveries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            item TEXT,
            scale TEXT,
            amount TEXT,
            period TEXT,
            client TEXT,
            project_name TEXT,
            note TEXT,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS product_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            changed_by INTEGER NOT NULL,
            changed_at TEXT NOT NULL,
            changed_fields TEXT,
            before_json TEXT,
            after_json TEXT,
            reason TEXT,
            FOREIGN KEY(product_id) REFERENCES products(id)
        );
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            action TEXT NOT NULL,
            detail TEXT,
            created_at TEXT NOT NULL
        );
        ''')

        admin = conn.execute(
            "SELECT id FROM users WHERE role='admin'"
        ).fetchone()

        if not admin:
            admin_company_name = os.environ.get('ADMIN_COMPANY_NAME')
            admin_password = os.environ.get('ADMIN_PASSWORD')
            admin_email = os.environ.get(
                'ADMIN_EMAIL',
                'admin@example.com'
            )

            if not admin_company_name or not admin_password:
                raise RuntimeError(
                    'ADMIN_COMPANY_NAME 또는 ADMIN_PASSWORD 환경변수가 설정되지 않았습니다.'
                )

            conn.execute(
                '''
                INSERT INTO users(
                    role,
                    business_no,
                    company_name,
                    contact_name,
                    phone,
                    email,
                    password_hash,
                    privacy_consent,
                    created_at
                )
                VALUES('admin', NULL, ?, ?, ?, ?, ?, 1, ?)
                ''',
                (
                    admin_company_name,
                    '관리자',
                    '',
                    admin_email,
                    generate_password_hash(admin_password),
                    datetime.now().isoformat(timespec='seconds')
                )
            )

        conn.commit()
    finally:
        conn.close()


def log_action(action, detail=''):
    conn = db()
    conn.execute('INSERT INTO audit_logs(user_id,action,detail,created_at) VALUES(?,?,?,?)',
                 (session.get('user_id'), action, detail, datetime.now().isoformat(timespec='seconds')))
    conn.commit(); conn.close()


def login_required(role=None):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not session.get('user_id'):
                flash('로그인이 필요합니다.')
                return redirect(url_for('login'))
            if role and session.get('role') != role:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return deco



@app.context_processor
def inject_globals():
    return dict(CATEGORIES=CATEGORIES, COMPANY_CERTS=COMPANY_CERTS, POLICY_PRODUCTS=POLICY_PRODUCTS, CERT_TYPES=CERT_TYPES, format_company_certs=format_company_certs)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/register', methods=['GET','POST'])
def register():
    if request.method == 'POST':
        business_no = ''.join(c for c in request.form['business_no'] if c.isdigit())
        company_name = request.form['company_name'].strip()
        if not request.form.get('privacy_consent'):
            flash('개인정보 수집·이용에 동의해야 가입할 수 있습니다.')
            return redirect(url_for('register'))
        certs = request.form.getlist('company_certs')
        conn = db()
        exists = conn.execute('SELECT 1 FROM users WHERE business_no=? OR lower(trim(company_name))=lower(trim(?))', (business_no, company_name)).fetchone()
        if exists:
            conn.close(); flash('이미 등록된 업체명 또는 사업자등록번호입니다.')
            return redirect(url_for('register'))
        conn.execute('''INSERT INTO users(business_no,company_name,company_certs,company_cert_other,contact_name,phone,email,password_hash,privacy_consent,created_at)
                        VALUES(?,?,?,?,?,?,?,?,1,?)''',
                     (business_no, company_name, json.dumps(certs, ensure_ascii=False), request.form.get('company_cert_other','').strip(),
                      request.form['contact_name'].strip(), request.form['phone'].strip(), request.form['email'].strip(),
                      generate_password_hash(request.form['password']), datetime.now().isoformat(timespec='seconds')))
        conn.commit(); conn.close()
        flash('회원가입이 완료되었습니다. 로그인해 주세요.')
        return redirect(url_for('login'))
    return render_template('register.html')


@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        conn = db()
        user = conn.execute('SELECT * FROM users WHERE company_name=? AND active=1', (request.form['company_name'].strip(),)).fetchone()
        if not user or not check_password_hash(user['password_hash'], request.form['password']):
            conn.close(); flash('업체명 또는 비밀번호가 올바르지 않습니다.')
            return redirect(url_for('login'))
        session.clear(); session['user_id']=user['id']; session['role']=user['role']; session['company_name']=user['company_name']
        conn.execute('UPDATE users SET last_login=? WHERE id=?', (datetime.now().isoformat(timespec='seconds'), user['id']))
        conn.commit(); conn.close(); log_action('login')
        return redirect(url_for('admin_dashboard' if user['role']=='admin' else 'vendor_dashboard'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    log_action('logout'); session.clear(); return redirect(url_for('index'))


@app.route('/vendor')
@login_required('vendor')
def vendor_dashboard():
    conn=db(); products=conn.execute('SELECT * FROM products WHERE user_id=? ORDER BY created_at DESC',(session['user_id'],)).fetchall(); conn.close()
    return render_template('vendor_dashboard.html', products=products)


def product_payload(form):
    return {
        'product_name': form['product_name'].strip(), 'category': form['category'], 'category_other': form.get('category_other','').strip(),
        'manufacturer': form['manufacturer'].strip(), 'model_name': form.get('model_name','').strip(), 'country': form['country'].strip(),
        'policy_products': json.dumps(form.getlist('policy_products'), ensure_ascii=False), 'policy_other': form.get('policy_other','').strip(),
        'size_spec': form.get('size_spec','').strip(), 'performance': form.get('performance','').strip(), 'lifespan': form.get('lifespan','').strip(),
        'applicable_facilities': form.get('applicable_facilities','').strip(), 'compatibility': form.get('compatibility','').strip(),
        'delivery_period': form.get('delivery_period','').strip()
    }


@app.route('/vendor/product/new', methods=['GET','POST'])
@login_required('vendor')
def product_new():
    conn=db(); user=conn.execute('SELECT * FROM users WHERE id=?',(session['user_id'],)).fetchone()
    if request.method=='POST':
        required_checks=['truth_confirm','reference_confirm','no_visit_confirm']
        if not all(request.form.get(x) for x in required_checks):
            conn.close(); flash('필수 확인사항에 모두 동의해야 등록할 수 있습니다.'); return redirect(request.url)
        p=product_payload(request.form); p['spec_file']=save_file(request.files.get('spec_file'))
        now=datetime.now().isoformat(timespec='seconds')
        cur=conn.execute('''INSERT INTO products(user_id,product_name,category,category_other,manufacturer,model_name,country,policy_products,policy_other,size_spec,performance,lifespan,applicable_facilities,compatibility,delivery_period,spec_file,created_at,updated_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                         (session['user_id'],p['product_name'],p['category'],p['category_other'],p['manufacturer'],p['model_name'],p['country'],p['policy_products'],p['policy_other'],p['size_spec'],p['performance'],p['lifespan'],p['applicable_facilities'],p['compatibility'],p['delivery_period'],p['spec_file'],now,now))
        product_id=cur.lastrowid
        save_children(conn, product_id, request)
        conn.commit(); conn.close(); log_action('product_create', str(product_id)); flash('제품이 등록되었습니다.')
        return redirect(url_for('vendor_dashboard'))
    conn.close(); return render_template('product_form.html', user=user, product=None)


def _value_at(values, index, default=''):
    """
    리스트에서 지정한 순번의 값을 안전하게 가져옵니다.
    값이 없으면 default를 반환합니다.
    """
    return values[index] if index < len(values) else default


def save_children(conn, product_id, req):
    """
    제품에 딸린 인증정보와 주요 납품실적을 저장합니다.

    수정 화면에서 기존 인증정보와 납품실적도 다시 전송되며,
    삭제 표시된 행만 저장하지 않습니다.

    기존 인증서 파일을 새 파일로 교체하지 않은 경우,
    기존 파일명을 그대로 유지합니다.
    """

    # ─────────────────────────────
    # 인증정보 저장
    # ─────────────────────────────

    cert_types = req.form.getlist('cert_type[]')
    cert_names = req.form.getlist('cert_name[]')
    agencies = req.form.getlist('agency[]')
    cert_nos = req.form.getlist('cert_no[]')
    issue_dates = req.form.getlist('issue_date[]')
    valid_froms = req.form.getlist('valid_from[]')
    valid_tos = req.form.getlist('valid_to[]')
    cert_notes = req.form.getlist('cert_note[]')

    # 기존 인증서 파일명
    existing_files = req.form.getlist('cert_existing_file[]')

    # 삭제 여부: 0이면 유지, 1이면 삭제
    cert_deletes = req.form.getlist('cert_delete[]')

    # 새로 첨부한 인증서 파일
    cert_files = req.files.getlist('cert_file[]')

    for i, cert_type in enumerate(cert_types):

        # 삭제를 선택한 인증정보는 다시 저장하지 않음
        if _value_at(cert_deletes, i, '0') == '1':
            continue

        # 인증대분류가 비어 있는 빈 행은 저장하지 않음
        if not cert_type.strip():
            continue

        uploaded_file = _value_at(cert_files, i, None)

        # 새 인증서 파일이 있으면 새 파일 저장
        if uploaded_file and uploaded_file.filename:
            file_name = save_file(uploaded_file)

        # 새 파일이 없으면 기존 파일명 유지
        else:
            file_name = _value_at(existing_files, i, '') or None

        conn.execute(
            '''
            INSERT INTO certifications(
                product_id,
                cert_type,
                cert_name,
                agency,
                cert_no,
                issue_date,
                valid_from,
                valid_to,
                no_expiry,
                file_path,
                note
            )
            VALUES(?,?,?,?,?,?,?,?,?,?,?)
            ''',
            (
                product_id,
                cert_type,
                _value_at(cert_names, i),
                _value_at(agencies, i),
                _value_at(cert_nos, i),
                _value_at(issue_dates, i),
                _value_at(valid_froms, i),
                _value_at(valid_tos, i),
                0,
                file_name,
                _value_at(cert_notes, i)
            )
        )

    # ─────────────────────────────
    # 주요 납품실적 저장
    # ─────────────────────────────

    items = req.form.getlist('delivery_item[]')
    scales = req.form.getlist('delivery_scale[]')
    amounts = req.form.getlist('delivery_amount[]')
    periods = req.form.getlist('delivery_period[]')
    clients = req.form.getlist('delivery_client[]')
    projects = req.form.getlist('delivery_project[]')
    delivery_notes = req.form.getlist('delivery_note[]')

    # 삭제 여부: 0이면 유지, 1이면 삭제
    delivery_deletes = req.form.getlist('delivery_delete[]')

    for i, item in enumerate(items):

        # 삭제를 선택한 납품실적은 다시 저장하지 않음
        if _value_at(delivery_deletes, i, '0') == '1':
            continue

        # 품목이 비어 있는 빈 행은 저장하지 않음
        if not item.strip():
            continue

        conn.execute(
            '''
            INSERT INTO deliveries(
                product_id,
                item,
                scale,
                amount,
                period,
                client,
                project_name,
                note
            )
            VALUES(?,?,?,?,?,?,?,?)
            ''',
            (
                product_id,
                item,
                _value_at(scales, i),
                _value_at(amounts, i),
                _value_at(periods, i),
                _value_at(clients, i),
                _value_at(projects, i),
                _value_at(delivery_notes, i)
            )
        )


@app.route('/vendor/product/<int:pid>/edit', methods=['GET', 'POST'])
@login_required('vendor')
def product_edit(pid):
    conn = db()

    # 로그인한 업체 본인의 제품만 조회
    product = conn.execute(
        '''
        SELECT *
        FROM products
        WHERE id=? AND user_id=?
        ''',
        (pid, session['user_id'])
    ).fetchone()

    if not product:
        conn.close()
        abort(404)

    user = conn.execute(
        'SELECT * FROM users WHERE id=?',
        (session['user_id'],)
    ).fetchone()

    # 기존 인증정보 조회
    certs = conn.execute(
        'SELECT * FROM certifications WHERE product_id=? ORDER BY id',
        (pid,)
    ).fetchall()

    # 기존 납품실적 조회
    deliveries = conn.execute(
        'SELECT * FROM deliveries WHERE product_id=? ORDER BY id',
        (pid,)
    ).fetchall()

    if request.method == 'POST':
        before = dict(product)
        p = product_payload(request.form)

        # 제품 사양서 새 파일이 없으면 기존 파일 유지
        new_file = save_file(request.files.get('spec_file'))
        p['spec_file'] = new_file or product['spec_file']

        now = datetime.now().isoformat(timespec='seconds')

        conn.execute(
            '''
            UPDATE products
            SET
                product_name=?,
                category=?,
                category_other=?,
                manufacturer=?,
                model_name=?,
                country=?,
                policy_products=?,
                policy_other=?,
                size_spec=?,
                performance=?,
                lifespan=?,
                applicable_facilities=?,
                compatibility=?,
                delivery_period=?,
                spec_file=?,
                updated_at=?
            WHERE id=?
            ''',
            (
                p['product_name'],
                p['category'],
                p['category_other'],
                p['manufacturer'],
                p['model_name'],
                p['country'],
                p['policy_products'],
                p['policy_other'],
                p['size_spec'],
                p['performance'],
                p['lifespan'],
                p['applicable_facilities'],
                p['compatibility'],
                p['delivery_period'],
                p['spec_file'],
                now,
                pid
            )
        )

        # 기존 자식 데이터를 먼저 삭제
        conn.execute(
            'DELETE FROM certifications WHERE product_id=?',
            (pid,)
        )

        conn.execute(
            'DELETE FROM deliveries WHERE product_id=?',
            (pid,)
        )

        # 수정화면에서 전달된 내용으로 다시 저장
        # 삭제 체크된 건은 save_children에서 제외됨
        save_children(conn, pid, request)

        after = {
            **p,
            'updated_at': now
        }

        changed = [
            key for key in p
            if str(before.get(key, '')) != str(p.get(key, ''))
        ]

        conn.execute(
            '''
            INSERT INTO product_history(
                product_id,
                changed_by,
                changed_at,
                changed_fields,
                before_json,
                after_json,
                reason
            )
            VALUES(?,?,?,?,?,?,?)
            ''',
            (
                pid,
                session['user_id'],
                now,
                ', '.join(changed),
                json.dumps(before, ensure_ascii=False),
                json.dumps(after, ensure_ascii=False),
                request.form.get('change_reason', '').strip()
            )
        )

        conn.commit()
        conn.close()

        log_action('product_edit', str(pid))
        flash('제품정보가 수정되었습니다.')

        return redirect(url_for('vendor_dashboard'))

    conn.close()

    # certs와 deliveries를 수정화면으로 전달해야
    # 기존 인증·납품실적이 화면에 표시됩니다.
    return render_template(
        'product_form.html',
        user=user,
        product=product,
        certs=certs,
        deliveries=deliveries
    )


@app.route('/vendor/product/<int:pid>/history')
@login_required('vendor')
def product_history(pid):
    conn=db(); owned=conn.execute('SELECT 1 FROM products WHERE id=? AND user_id=?',(pid,session['user_id'])).fetchone()
    if not owned: conn.close(); abort(404)
    rows=conn.execute('SELECT * FROM product_history WHERE product_id=? ORDER BY changed_at DESC',(pid,)).fetchall(); conn.close()
    return render_template('history.html',rows=rows)


@app.route('/admin')
@login_required('admin')
def admin_dashboard():
    conn=db()
    vendor_count=conn.execute("SELECT COUNT(*) c FROM users WHERE role='vendor'").fetchone()['c']
    product_count=conn.execute('SELECT COUNT(*) c FROM products').fetchone()['c']
    recent=conn.execute('''SELECT p.*,u.company_name FROM products p JOIN users u ON p.user_id=u.id ORDER BY p.created_at DESC LIMIT 10''').fetchall()
    conn.close(); return render_template('admin_dashboard.html',vendor_count=vendor_count,product_count=product_count,recent=recent)


@app.route('/admin/vendors')
@login_required('admin')
def admin_vendors():
    conn = db()
    rows = conn.execute('''
        SELECT
            u.*,
            COUNT(p.id) AS product_count,
            GROUP_CONCAT(DISTINCT
                CASE
                    WHEN p.category = '기타' AND COALESCE(p.category_other, '') <> ''
                        THEN '기타(' || p.category_other || ')'
                    ELSE p.category
                END
            ) AS product_categories
        FROM users u
        LEFT JOIN products p ON u.id = p.user_id
        WHERE u.role = 'vendor'
        GROUP BY u.id
        ORDER BY u.created_at DESC
    ''').fetchall()
    conn.close()

    return render_template(
        'admin_vendors.html',
        rows=rows
    )


@app.route('/admin/vendors/export.xlsx')
@login_required('admin')
def export_vendors_xlsx():
    conn = db()
    rows = conn.execute('''
        SELECT
            u.*,
            COUNT(p.id) AS product_count,
            GROUP_CONCAT(DISTINCT
                CASE
                    WHEN p.category = '기타' AND COALESCE(p.category_other, '') <> ''
                        THEN '기타(' || p.category_other || ')'
                    ELSE p.category
                END
            ) AS product_categories
        FROM users u
        LEFT JOIN products p ON u.id = p.user_id
        WHERE u.role = 'vendor'
        GROUP BY u.id
        ORDER BY u.created_at DESC
    ''').fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = '등록업체'
    headers = ['업체명','사업자등록번호','제품 카테고리','담당자명','연락처','이메일','가입일','최근 로그인','등록제품 수','회원상태']
    ws.append(headers)
    for r in rows:
        ws.append([
            r['company_name'], r['business_no'],
            r['product_categories'] or '등록제품 없음',
            r['contact_name'], r['phone'], r['email'], r['created_at'],
            r['last_login'] or '', r['product_count'], '활성' if r['active'] else '비활성'
        ])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col) + 2, 40)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    conn.close()
    log_action('vendor_excel_download', '등록업체 전체')
    return send_file(
        out,
        as_attachment=True,
        download_name='등록업체현황.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/admin/vendor/<int:uid>')
@login_required('admin')
def admin_vendor_detail(uid):
    conn=db(); user=conn.execute("SELECT * FROM users WHERE id=? AND role='vendor'",(uid,)).fetchone(); products=conn.execute('SELECT * FROM products WHERE user_id=? ORDER BY created_at DESC',(uid,)).fetchall(); conn.close()
    if not user: abort(404)
    return render_template('admin_vendor_detail.html',user=user,products=products)


@app.route('/admin/products')
@login_required('admin')
def admin_products():
    cat=request.args.get('category',''); q=request.args.get('q','').strip(); order=request.args.get('order','new')
    sql='''SELECT p.*,u.company_name,u.contact_name,u.phone,u.email FROM products p JOIN users u ON p.user_id=u.id WHERE 1=1'''; params=[]
    if cat: sql+=' AND p.category=?'; params.append(cat)
    if q: sql+=' AND (p.product_name LIKE ? OR u.company_name LIKE ? OR p.manufacturer LIKE ?)'; params += [f'%{q}%']*3
    sql += ' ORDER BY p.created_at ' + ('ASC' if order=='old' else 'DESC')
    conn=db(); rows=conn.execute(sql,params).fetchall(); conn.close()
    return render_template('admin_products.html',rows=rows,category=cat,q=q,order=order)


@app.route('/admin/product/<int:pid>')
@login_required('admin')
def admin_product_detail(pid):
    conn=db(); product=conn.execute('''SELECT p.*,u.company_name,u.business_no,u.contact_name,u.phone,u.email FROM products p JOIN users u ON p.user_id=u.id WHERE p.id=?''',(pid,)).fetchone()
    certs=conn.execute('SELECT * FROM certifications WHERE product_id=?',(pid,)).fetchall(); deliveries=conn.execute('SELECT * FROM deliveries WHERE product_id=?',(pid,)).fetchall(); history=conn.execute('SELECT * FROM product_history WHERE product_id=? ORDER BY changed_at DESC',(pid,)).fetchall(); conn.close()
    if not product: abort(404)
    log_action('product_view',str(pid)); return render_template('admin_product_detail.html',product=product,certs=certs,deliveries=deliveries,history=history)


@app.route('/admin/export.xlsx')
@login_required('admin')
def export_xlsx():
    cat=request.args.get('category',''); sql='''SELECT u.company_name,u.business_no,u.contact_name,u.phone,u.email,p.* FROM products p JOIN users u ON p.user_id=u.id WHERE 1=1'''; params=[]
    if cat: sql+=' AND p.category=?'; params.append(cat)
    sql+=' ORDER BY p.created_at DESC'
    conn=db(); rows=conn.execute(sql,params).fetchall()
    wb=Workbook(); ws=wb.active; ws.title='등록제품'
    headers=['업체명','사업자등록번호','담당자명','연락처','이메일','제품명','카테고리','제조사','모델명','제조국','정부권장정책제품','규격 및 크기','주요 성능','예상 내용연수','적용 가능한 시설','호환조건','납품·설치 소요기간','최초 등록일','최종 수정일']
    ws.append(headers)
    for r in rows:
        ws.append([r['company_name'],r['business_no'],r['contact_name'],r['phone'],r['email'],r['product_name'],r['category'],r['manufacturer'],r['model_name'],r['country'],', '.join(json.loads(r['policy_products'] or '[]')),r['size_spec'],r['performance'],r['lifespan'],r['applicable_facilities'],r['compatibility'],r['delivery_period'],r['created_at'],r['updated_at']])
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width=min(max(len(str(c.value or '')) for c in col)+2,35)
    out=BytesIO(); wb.save(out); out.seek(0); conn.close(); log_action('excel_download',cat or '전체')
    return send_file(out,as_attachment=True,download_name='등록제품현황.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/uploads/<path:name>')
@login_required()
def uploaded_file(name):
    # 업체는 자신이 등록한 첨부파일만, 관리자는 전체 열람 가능
    if session.get('role') == 'vendor':
        conn = db()

        own = conn.execute(
            '''
            SELECT 1
            FROM products
            WHERE user_id=? AND spec_file=?
            ''',
            (session['user_id'], name)
        ).fetchone()

        if not own:
            own = conn.execute(
                '''
                SELECT 1
                FROM certifications c
                JOIN products p ON c.product_id=p.id
                WHERE p.user_id=? AND c.file_path=?
                ''',
                (session['user_id'], name)
            ).fetchone()

        conn.close()

        if not own:
            abort(403)

    return send_from_directory(
        UPLOAD_DIR,
        name,
        as_attachment=True
    )


if __name__ == '__main__':
    init_db()
    app.run(
        host='0.0.0.0',
        port=int(os.environ.get('PORT', 5000)),
        debug=False
    )
