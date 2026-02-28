# ================= IMPORTAÇÕES =================
from flask import (
    Flask,
    render_template,
    jsonify,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file
)

from openpyxl import load_workbook, Workbook
from functools import wraps
import os
import calendar
import datetime

# ================= CONFIG FLASK =================
app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'NWanClh3BDY8I67SwHmXjhPQ2We2n2GMbr7KOtRIeJ7s9KMOMp'

app.config.update(
    SESSION_PERMANENT=False,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=False
)

# ================= CONFIG EXCEL =================
ARQUIVO_SIG = 'dados.xlsx'
ARQUIVO_SSH = 'dadossh.xlsx'
ANO_FIXO = 2026

# Cache (evita lentidão e crash)
cache_total_geral = {"dados": None}

# ================= HELPERS =================
def garantir_arquivo(arquivo):
    """Cria o arquivo Excel se não existir (ESSENCIAL no Render)"""
    if not os.path.exists(arquivo):
        wb = Workbook()
        wb.save(arquivo)

# 🔥 MUITO IMPORTANTE PARA RENDER (evita 502)
garantir_arquivo(ARQUIVO_SIG)
garantir_arquivo(ARQUIVO_SSH)

def garantir_aba(arquivo, mes, tipo):
    garantir_arquivo(arquivo)
    mes = mes.upper()

    if mes == 'TOTAL GERAL':
        return

    wb = load_workbook(arquivo)

    if mes not in wb.sheetnames:
        ws = wb.create_sheet(mes)

        ws['A1'] = 'ID'
        ws['B1'] = 'DATA'
        ws['C1'] = 'P&R'

        if tipo == 'sig':
            ws['D1'] = 'EMBAIXADOR'
            ws['F1'] = 'CSS'
            ws['G1'] = '% CSS'

        meses = {
            'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'ABRIL':4,
            'MAIO':5,'JUNHO':6,'JULHO':7,'AGOSTO':8,
            'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12
        }

        numero = meses.get(mes, 1)
        ultimo = calendar.monthrange(ANO_FIXO, numero)[1]

        for d in range(1, ultimo + 1):
            data = datetime.date(ANO_FIXO, numero, d)
            ws.cell(row=d+1, column=1, value=d)
            ws.cell(row=d+1, column=2, value=data.strftime('%d/%m/%Y'))

        wb.save(arquivo)

# ================= NO CACHE HEADERS =================
@app.after_request
def no_cache(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

# ================= USUÁRIOS =================
USUARIOS = {'admin': 'sig@2025'}

# ================= LOGIN =================
@app.route('/Login-Planilha', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session.clear()
        usuario = request.form.get('usuario')
        senha = request.form.get('senha')

        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            session['usuario'] = usuario
            session.permanent = False
            return redirect(url_for('planilha_sig'))

        flash('Usuário ou senha incorretos')

    return render_template('login.html')

# ================= MIDDLEWARE =================
def login_required(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrap

# ================= ROTAS PRINCIPAIS =================
@app.route('/')
def index():
    if 'usuario' in session:
        return redirect(url_for('planilha_sig'))
    return redirect(url_for('login'))

@app.route('/Home')
@login_required
def home():
    return render_template('inicio.html')

@app.route('/planilha-sig')
@login_required
def planilha_sig():
    return render_template('index.html', tipo='sig')

@app.route('/planilha-ssh')
@login_required
def planilha_ssh():
    return render_template('index.html', tipo='ssh')

# ================= API MESES =================
@app.route('/api/meses')
@login_required
def api_meses():
    garantir_arquivo(ARQUIVO_SIG)
    wb = load_workbook(ARQUIVO_SIG, read_only=True, data_only=True)

    MESES_ORDEM = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO',
        'TOTAL GERAL'
    ]

    abas = set(s.strip().upper() for s in wb.sheetnames)
    resultado = [mes for mes in MESES_ORDEM if mes in abas]

    return jsonify(resultado)

# ================= API SALVAR =================
@app.route('/api/salvar', methods=['POST'])
@login_required
def api_salvar():
    try:
        data = request.json
        mes = data.get('mes')
        dia = int(data.get('dia'))
        pr = data.get('pr')
        emb = data.get('emb')
        css = data.get('css')
        tipo = data.get('tipo')
        percent_css = data.get('percent_css')

        if mes and mes.upper() == 'TOTAL GERAL':
            return jsonify({'error': 'TOTAL GERAL não pode ser editado'}), 403

        arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH
        garantir_aba(arquivo, mes, tipo)

        wb = load_workbook(arquivo)
        ws = wb[mes.upper()]

        if pr not in (None, ''):
            ws.cell(row=dia+1, column=3, value=float(str(pr).replace(',', '.')))

        ws.cell(row=dia+1, column=4, value=emb if emb else '')

        if css not in (None, ''):
            ws.cell(row=dia+1, column=6, value=float(str(css).replace(',', '.')))

        if percent_css not in (None, ''):
            ws.cell(row=dia+1, column=7, value=float(str(percent_css).replace(',', '.')))

        wb.save(arquivo)
        cache_total_geral["dados"] = None  # limpa cache

        return jsonify({'ok': True})

    except Exception as e:
        print("ERRO AO SALVAR:", str(e))
        return jsonify({'error': str(e)}), 500

# ================= API TABELA =================
@app.route('/api/tabela')
@login_required
def api_tabela():
    mes = request.args.get('mes')
    tipo = request.args.get('tipo')

    arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH
    if not os.path.exists(arquivo):
        return jsonify([])

    wb = load_workbook(arquivo, data_only=True)
    if mes.upper() not in wb.sheetnames:
        return jsonify([])

    ws = wb[mes.upper()]
    dados = []

    for r in range(2, ws.max_row + 1):
        data_celula = ws.cell(row=r, column=2).value

        if isinstance(data_celula, (datetime.date, datetime.datetime)):
            data_formatada = data_celula.strftime('%d/%m/%Y')
        else:
            data_formatada = data_celula or ''

        item = {
            'id': ws.cell(row=r, column=1).value or '',
            'data': data_formatada,
            'pr': ws.cell(row=r, column=3).value or ''
        }

        if tipo == 'sig':
            item['emb'] = ws.cell(row=r, column=4).value or ''
            item['css'] = ws.cell(row=r, column=6).value or ''
            item['percent_css'] = ws.cell(row=r, column=7).value or ''

        dados.append(item)

    return jsonify(dados)

# ================= DOWNLOAD =================
@app.route("/baixar-sig")
@login_required
def baixar_sig():
    if os.path.exists(ARQUIVO_SIG):
        return send_file(ARQUIVO_SIG, as_attachment=True)
    return "Arquivo dados.xlsx não encontrado", 404

@app.route("/baixar-ssh")
@login_required
def baixar_ssh():
    if os.path.exists(ARQUIVO_SSH):
        return send_file(ARQUIVO_SSH, as_attachment=True)
    return "Arquivo dadossh.xlsx não encontrado", 404

# ================= RUN (RENDER SAFE) =================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=False)
