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
    send_file  # 🔥 ADICIONE ISSO
)


from openpyxl import load_workbook, Workbook
from functools import wraps
import os
import calendar
import datetime
import send_file

# ================= CONFIG FLASK =================
app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'NWanClh3BDY8I67SwHmXjhPQ2We2n2GMbr7KOtRIeJ7s9KMOMp'

# 🔒 CONFIGURAÇÕES PARA NÃO SALVAR LOGIN
app.config.update(
    SESSION_PERMANENT=False,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='None',
    SESSION_COOKIE_SECURE=True,  # 🔥 OBRIGATÓRIO no Render
)

@app.after_request
def no_cache(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


# ================= USUÁRIOS =================
USUARIOS = {'admin': 'sig@2025'}

# ================= CONFIG EXCEL =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Se existir disco persistente do Render usa /data
if os.path.exists('/data'):
    DATA_DIR = '/data'
else:
    DATA_DIR = BASE_DIR  # funciona local também

ARQUIVO_SIG = os.path.join(DATA_DIR, 'dados.xlsx')
ARQUIVO_SSH = os.path.join(DATA_DIR, 'dadossh.xlsx')
ANO_FIXO = 2026

# ================= LOGIN =================
@app.route('/Login-Planilha', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session.clear()  # 🔥 limpa sessão APENAS ao tentar logar

        usuario = request.form.get('usuario')
        senha = request.form.get('senha')

        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            session['usuario'] = usuario
            session.permanent = False  # 🔒 não salva login
            return redirect(url_for('planilha_sig'))

        flash('Usuário ou senha incorretos')

    return render_template('login.html')



# ================= MIDDLEWARE =================
def login_required(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'usuario' not in session:
            # 🔥 SE FOR API, NÃO REDIRECIONA (isso evita o loop infinito)
            if request.path.startswith('/api/'):
                return jsonify({'error': 'nao_autorizado'}), 401
            
            # 🔒 Se for página normal, redireciona para login
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrap

# ================= HELPERS =================
def garantir_arquivo(arquivo):
    if not os.path.exists(arquivo):
        wb = Workbook()
        wb.save(arquivo)

def garantir_aba(arquivo, mes, tipo):
    garantir_arquivo(arquivo)
    mes = mes.upper()

    wb = load_workbook(arquivo)  # ✅ wb definido aqui

    # 🚫 TOTAL GERAL nunca é criado nem alterado
    if mes == 'TOTAL GERAL':
        return

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


# ================= ROTAS PÚBLICAS =================
@app.route('/Home')
@login_required  # 🔒 AGORA EXIGE LOGIN
def home():
    return render_template('inicio.html')


# ================= PLANILHAS (CADA UMA É UM "SITE") =================
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

    wb = load_workbook(ARQUIVO_SIG, read_only=True)

    MESES_ORDEM = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO',
        'TOTAL GERAL'
    ]

    abas = set(s.strip().upper() for s in wb.sheetnames)

    resultado = [mes for mes in MESES_ORDEM if mes in abas]

    return jsonify(resultado)

# ================= BAIXAR PLANILHA SIG =================
@app.route('/baixar-sig')
@login_required
def baixar_sig():
    try:
        if not os.path.exists(ARQUIVO_SIG):
            return "Arquivo SIG não encontrado", 404

        return send_file(
            ARQUIVO_SIG,
            as_attachment=True,
            download_name='planilha_SIG_atualizada.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            cache_timeout=0  # 🔥 evita baixar versão antiga
        )
    except Exception as e:
        return str(e), 500


# ================= BAIXAR PLANILHA SSH =================
@app.route('/baixar-ssh')
@login_required
def baixar_ssh():
    try:
        if not os.path.exists(ARQUIVO_SSH):
            return "Arquivo SSH não encontrado", 404

        return send_file(
            ARQUIVO_SSH,
            as_attachment=True,
            download_name='planilha_SSH_atualizada.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            cache_timeout=0
        )
    except Exception as e:
        return str(e), 500


# ================= API DIAS =================
@app.route('/api/dias')
@login_required
def api_dias():
    mes = request.args.get('mes')
    if not mes:
        return jsonify([])

    meses = {
        'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'ABRIL':4,
        'MAIO':5,'JUNHO':6,'JULHO':7,'AGOSTO':8,
        'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12
    }

    numero = meses.get(mes.upper(), 1)
    ultimo = calendar.monthrange(ANO_FIXO, numero)[1]
    return jsonify(list(range(1, ultimo + 1)))

# ================= ATUALIZAR TOTAL GERAL NO EXCEL =================
def atualizar_total_geral_excel():
    MESES_VALIDOS = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'
    ]

    # 🚀 MODO RÁPIDO
    wb = load_workbook(ARQUIVO_SIG, read_only=True, data_only=True)

    totais = []
    total_pr_anual = 0
    total_css_anual = 0
    soma_css_peso_anual = 0

    for mes in MESES_VALIDOS:
        if mes not in wb.sheetnames:
            continue

        ws = wb[mes]

        total_pr_mes = 0
        soma_css_mes = 0
        soma_css_peso_mes = 0

        # ⚡ MUITO MAIS RÁPIDO que ws.cell()
        for row in ws.iter_rows(min_row=2, values_only=True):
            pr = row[2]   # Coluna C
            css = row[5]  # Coluna F
            percent = row[6]  # Coluna G

            if pr not in (None, ''):
                try:
                    total_pr_mes += float(pr)
                except:
                    pass

            if css not in (None, '') and percent not in (None, ''):
                try:
                    css = float(css)
                    percent = float(percent)

                    if css > 0:
                        soma_css_mes += css
                        soma_css_peso_mes += css * percent
                except:
                    pass

        media_percent = (
            round(soma_css_peso_mes / soma_css_mes, 1)
            if soma_css_mes > 0 else 0
        )

        totais.append({
            'mes': mes,
            'pr': int(total_pr_mes),
            'css': int(soma_css_mes),
            'percent': media_percent
        })

        total_pr_anual += total_pr_mes
        total_css_anual += soma_css_mes
        soma_css_peso_anual += soma_css_peso_mes

    media_anual = (
        round(soma_css_peso_anual / total_css_anual, 1)
        if total_css_anual > 0 else 0
    )

    return totais, int(total_pr_anual), int(total_css_anual), media_anual


    # 🔥 TOTAL ANUAL (ÚLTIMA LINHA)
    ws.cell(row=linha, column=1, value='TOTAL ANUAL')
    ws.cell(row=linha, column=3, value=int(total_pr_anual))
    ws.cell(row=linha, column=6, value=int(total_css_anual))

    wb.save(ARQUIVO_SIG)


# ================= API SALVAR =================
@app.route('/api/salvar', methods=['POST'])
@login_required
def api_salvar():
    try:
        data = request.json

        mes = (data.get('mes') or '').strip().upper()
        dia = int(data.get('dia'))
        pr = data.get('pr')
        emb = data.get('emb')
        css = data.get('css')
        percent_css = data.get('percent_css')
        tipo = data.get('tipo')

        if mes == 'TOTAL GERAL':
            return jsonify({'error': 'TOTAL GERAL não pode ser editado'}), 403

        arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH

        wb = load_workbook(arquivo)

        # 🔥 PEGA SOMENTE ABA SEM NÚMERO (EX: FEVEREIRO)
        aba_correta = None

        for sheet in wb.sheetnames:
            nome = sheet.strip()
            nome_upper = nome.upper()

            # tem que ser o mês
            if not nome_upper.startswith(mes):
                continue

            # ignora abas com número (MAIO5, JUNHO6, etc)
            if any(c.isdigit() for c in nome_upper):
                continue

            aba_correta = sheet
            break

        if not aba_correta:
            return jsonify({
                'error': f'Aba "{mes}" sem número não encontrada',
                'abas': wb.sheetnames
            }), 400

        ws = wb[aba_correta]

        linha = dia + 1  # linha 2 = dia 1

        # P&R → Coluna C (3)
        if pr not in (None, ''):
            ws.cell(row=linha, column=3,
                    value=float(str(pr).replace(',', '.')))

        # EMBAIXADOR → Coluna D (4)
        ws.cell(row=linha, column=4, value=emb if emb else '')

        # CSS → Coluna F (6)
        if css not in (None, ''):
            ws.cell(row=linha, column=6,
                    value=float(str(css).replace(',', '.')))

        # % CSS → Coluna G (7)
        if percent_css not in (None, ''):
            ws.cell(row=linha, column=7,
                    value=float(str(percent_css).replace(',', '.')))

        wb.save(arquivo)
        wb.close()

        print(f"[SALVO NA ABA SEM NÚMERO]: {aba_correta}")

        return jsonify({'ok': True, 'aba_usada': aba_correta})

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

        # ✅ FORMATA DATA CORRETAMENTE
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


# ================= RESUMO =================
@app.route('/resumo')
@login_required
def resumo():

    MESES_VALIDOS = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'
    ]

    def soma_coluna(arquivo, coluna):
        total = 0
        if not os.path.exists(arquivo):
            return 0

        wb = load_workbook(arquivo, data_only=True)

        for aba in wb.sheetnames:
            if aba.upper() not in MESES_VALIDOS:
                continue  # ignora abas inválidas

            ws = wb[aba]
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=coluna).value
                if v not in (None, ''):
                    try:
                        total += float(str(v).replace(',', '.'))
                    except:
                        pass

        return total

    # P&R
    total_sig_pr = soma_coluna(ARQUIVO_SIG, 3)
    total_ssh_pr = soma_coluna(ARQUIVO_SSH, 3)

    # CSS (somente SIG)
    total_sig_css = soma_coluna(ARQUIVO_SIG, 6)

    resultado = int(total_sig_pr - total_ssh_pr)

    return render_template(
        'resumo.html',
        total_sig=int(total_sig_pr),
        total_ssh=int(total_ssh_pr),
        total_css=int(total_sig_css),
        resultado=resultado
    )


# ================= TOTAL GERAL COMO "MÊS" =================
@app.route('/api/mes-total-geral')
@login_required
def api_mes_total_geral():

    tipo = request.args.get('tipo')
    arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH

    if not os.path.exists(arquivo):
        return jsonify([])

    MESES_ORDEM = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'
    ]

    try:
        wb = load_workbook(arquivo, read_only=True, data_only=True)
        resultado = []

        total_pr_anual = 0
        total_css_anual = 0
        soma_css_peso_anual = 0

        for mes in MESES_ORDEM:
            if mes not in wb.sheetnames:
                continue

            ws = wb[mes]

            total_pr_mes = 0
            soma_css_mes = 0
            soma_css_peso_mes = 0

            for row in ws.iter_rows(min_row=2, values_only=True):
                pr = row[2] if len(row) > 2 else None
                css = row[5] if len(row) > 5 else None
                percent = row[6] if len(row) > 6 else None

                # P&R (funciona para SIG e SSH)
                if pr not in (None, ''):
                    try:
                        total_pr_mes += float(pr)
                    except:
                        pass

                # 🔥 CSS só existe no SIG (evita erro no SSH)
                if tipo == 'sig' and css and percent:
                    try:
                        css = float(css)
                        percent = float(percent)

                        if css > 0:
                            soma_css_mes += css
                            soma_css_peso_mes += css * percent
                    except:
                        pass

            # Percentual só para SIG
            if tipo == 'sig' and soma_css_mes > 0:
                media_percent_mes = round(soma_css_peso_mes / soma_css_mes, 1)
            else:
                media_percent_mes = 0

            resultado.append({
                'id': '',
                'data': mes,
                'pr': int(total_pr_mes),
                'css': int(soma_css_mes) if tipo == 'sig' else 0,
                'percent_css': media_percent_mes
            })

            total_pr_anual += total_pr_mes
            if tipo == 'sig':
                total_css_anual += soma_css_mes
                soma_css_peso_anual += soma_css_peso_mes

        # TOTAL GERAL
        if tipo == 'sig' and total_css_anual > 0:
            media_anual = round(soma_css_peso_anual / total_css_anual, 1)
        else:
            media_anual = 0

        resultado.append({
            'id': '',
            'data': 'TOTAL GERAL',
            'pr': int(total_pr_anual),
            'css': int(total_css_anual) if tipo == 'sig' else 0,
            'percent_css': media_anual
        })

        return jsonify(resultado)

    except Exception as e:
        print("ERRO TOTAL GERAL:", str(e))
        return jsonify([]), 500


# ================= OUTRAS =================
@app.route('/calculadora')
@login_required
def calculadora():
    return render_template('calculadora.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ================= ROTA RAIZ =================
@app.route('/')
def index():
    # se não estiver logado, mostra login direto (SEM loop)
    if 'usuario' not in session:
        return render_template('login.html')

    # se estiver logado vai para planilha
    return redirect(url_for('planilha_sig'))


# ================= RUN =================
# NÃO coloque debug=True em produção
if __name__ == '__main__':
    app.run()



