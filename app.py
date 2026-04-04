# ================= IMPORTAÇÕES =================
from flask import (
    Flask,
    render_template,
    jsonify,
    request,
    redirect,
    url_for,
    session,
    flash
)

from openpyxl import load_workbook, Workbook
from functools import wraps
import os
import json
import calendar
import datetime
import requests
from bs4 import BeautifulSoup
import subprocess
import uuid
import os
from flask import request, jsonify, send_file

# ================= CONFIG FLASK =================
app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'NWanClh3BDY8I67SwHmXjhPQ2We2n2GMbr7KOtRIeJ7s9KMOMp'

# 🔒 CONFIGURAÇÕES PARA NÃO SALVAR LOGIN
app.config.update(
    SESSION_PERMANENT=False,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=False,   # True só se usar HTTPS
)

@app.after_request
def no_cache(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

import requests
import re
import json
import os
from bs4 import BeautifulSoup


def pegar_lancamentos():

    url = "https://www.mcdonalds.com.br/lancamentos"

    r = requests.get(url)

    soup = BeautifulSoup(r.text,"html.parser")

    lancamentos = []

    ano_atual = datetime.datetime.now().year

    for a in soup.find_all("a",href=True):

        if "/lancamentos/" in a["href"]:

            link = "https://www.mcdonalds.com.br" + a["href"]

            nome = a.get_text(strip=True)

            img = a.find("img")

            if img:
                imagem = img.get("src")
            else:
                imagem = ""

            if nome:

                lancamentos.append({

                    "nome": nome,
                    "link": link,
                    "imagem": imagem,
                    "ano": ano_atual

                })

    return lancamentos



def salvar_historico(lancamentos):

    arquivo = "lancamentos.json"

    if os.path.exists(arquivo):

        with open(arquivo,"r",encoding="utf-8") as f:
            historico = json.load(f)

    else:
        historico = []

    nomes = [i["nome"] for i in historico]

    for item in lancamentos:

        if item["nome"] not in nomes:

            historico.append(item)

    with open(arquivo,"w",encoding="utf-8") as f:

        json.dump(historico,f,indent=4,ensure_ascii=False)

    return historico
# ================= USUÁRIOS =================
USUARIOS = {'admin': 'sig@2025'}

# ================= CONFIG EXCEL =================
ARQUIVO_SIG = 'dados.xlsx'
ARQUIVO_SSH = 'dadossh.xlsx'
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
        mes = data.get('mes')
        dia = int(data.get('dia'))
        pr = data.get('pr')
        emb = data.get('emb')
        css = data.get('css')
        tipo = data.get('tipo')
        percent_css = data.get('percent_css')

        # 🚫 BLOQUEIO TOTAL
        if mes and mes.upper() == 'TOTAL GERAL':
            return jsonify({'error': 'TOTAL GERAL não pode ser editado'}), 403

        arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH
        garantir_aba(arquivo, mes, tipo)

        wb = load_workbook(arquivo)
        ws = wb[mes.upper()]

        # P&R → coluna C
        if pr not in (None, ''):
            ws.cell(
                row=dia+1,
                column=3,
                value=float(str(pr).replace(',', '.'))
            )

        # 🔥 EMBAIXADOR → coluna D (AGORA VAI SALVAR)
        ws.cell(
            row=dia+1,
            column=4,
            value=emb if emb else ''
        )

        # CSS → coluna F
        if css not in (None, ''):
            ws.cell(
                row=dia+1,
                column=6,
                value=float(str(css).replace(',', '.'))
            )

        # % CSS → coluna G
        if percent_css not in (None, ''):
            ws.cell(
                row=dia+1,
                column=7,
                value=float(str(percent_css).replace(',', '.'))
            )

        wb.save(arquivo)

        return jsonify({'ok': True})

    except Exception as e:
        print("ERRO AO SALVAR:", str(e))  # 🔥 vai aparecer no log do Render
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
    global cache_total_geral

    tipo = request.args.get('tipo')
    arquivo = ARQUIVO_SIG if tipo == 'sig' else ARQUIVO_SSH

    # ⚡ Se já tem cache, retorna instantâneo
    if cache_total_geral["dados"] is not None:
        return jsonify(cache_total_geral["dados"])

    if not os.path.exists(arquivo):
        return jsonify([])

    MESES_ORDEM = [
        'JANEIRO','FEVEREIRO','MARÇO','ABRIL',
        'MAIO','JUNHO','JULHO','AGOSTO',
        'SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'
    ]

    # 🚀 MODO RÁPIDO PARA RENDER
    wb = load_workbook(arquivo, read_only=True, data_only=True)
    resultado = []

    total_pr_anual = 0
    soma_css_peso_anual = 0
    soma_css_anual = 0

    for mes in MESES_ORDEM:
        if mes not in wb.sheetnames:
            continue

        ws = wb[mes]

        total_pr_mes = 0
        soma_css_mes = 0
        soma_css_peso_mes = 0

        # ⚡ 10x mais rápido que ws.cell()
        for row in ws.iter_rows(min_row=2, values_only=True):
            pr = row[2]
            css = row[5]
            percent = row[6]

            if pr:
                try:
                    total_pr_mes += float(pr)
                except:
                    pass

            if css and percent:
                try:
                    css = float(css)
                    percent = float(percent)

                    if css > 0:
                        soma_css_mes += css
                        soma_css_peso_mes += css * percent
                except:
                    pass

        media_percent_mes = (
            round(soma_css_peso_mes / soma_css_mes, 1)
            if soma_css_mes > 0 else 0
        )

        resultado.append({
            'id': '',
            'data': mes,
            'pr': int(total_pr_mes),
            'css': int(soma_css_mes),
            'percent_css': media_percent_mes
        })

        total_pr_anual += total_pr_mes
        soma_css_anual += soma_css_mes
        soma_css_peso_anual += soma_css_peso_mes

    media_percent_anual = (
        round(soma_css_peso_anual / soma_css_anual, 1)
        if soma_css_anual > 0 else 0
    )

    resultado.append({
        'id': '',
        'data': 'TOTAL GERAL',
        'pr': int(total_pr_anual),
        'css': int(soma_css_anual),
        'percent_css': media_percent_anual
    })

    # 🔥 SALVA NO CACHE (agora fica instantâneo)
    cache_total_geral["dados"] = resultado

    return jsonify(resultado)

PASTA_DOWNLOAD = "downloads"
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

# ================= OUTRAS =================
@app.route("/novidades")
def novidades():

    lancamentos = pegar_lancamentos()

    historico = salvar_historico(lancamentos)

    historico_por_ano = {}

    for item in historico:

        ano = item["ano"]

        if ano not in historico_por_ano:
            historico_por_ano[ano] = []

        historico_por_ano[ano].append(item)

    return render_template(

        "novidades.html",

        lancamentos=lancamentos,
        historico=historico_por_ano

    )

@app.route("/api/download", methods=["POST"])
def download_video():

    url = request.json.get("url")
    tipo = request.json.get("tipo")

    print("URL:", url)
    print("TIPO:", tipo)

    nome = str(uuid.uuid4())

    if tipo == "audio":

        caminho = os.path.join(
            PASTA_DOWNLOAD,
            f"{nome}.mp3"
        )

        comando = [
            "yt-dlp",
            "-x",
            "--audio-format",
            "mp3",
            "-o",
            caminho,
            url
        ]

    else:

        caminho = os.path.join(
            PASTA_DOWNLOAD,
            f"{nome}.mp4"
        )

        comando = [
            "yt-dlp",
            "-f",
            "mp4",
            "-o",
            caminho,
            url
        ]

    resultado = subprocess.run(
        comando,
        capture_output=True,
        text=True
    )

    print("STDOUT:", resultado.stdout)
    print("STDERR:", resultado.stderr)

    if resultado.returncode != 0:
        return jsonify({
            "erro": resultado.stderr
        }), 500

    if not os.path.exists(caminho):
        return jsonify({
            "erro": "Arquivo não foi criado"
        }), 500

    from flask import after_this_request

    @after_this_request
    def remover_arquivo(response):
        try:
            if os.path.exists(caminho):
                os.remove(caminho)
        except Exception as e:
            print("Erro ao remover arquivo:", e)
        return response

    return send_file(
     caminho,
    as_attachment=True,
    download_name=os.path.basename(caminho)

)

from flask import request, jsonify
import yt_dlp


@app.route("/api/info", methods=["POST"])
def info_video():

    try:

        data = request.get_json()

        url = data.get("url")

        if not url:

            return jsonify({
                "erro": "URL vazia"
            })

        ydl_opts = {
            "quiet": True,
            "skip_download": True
        }

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:

            info = ydl.extract_info(
                url,
                download=False
            )

            return jsonify({

                "titulo": info.get("title"),

                "thumbnail": info.get("thumbnail")

            })

    except Exception as e:

        print("ERRO INFO:", e)

        return jsonify({
            "erro": str(e)
        })

@app.route('/calculadora')
@login_required
def calculadora():
    return render_template('calculadora.html')

@app.route("/download")
def pagina_download():
    return render_template("download.html")

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ================= ROTA RAIZ =================
@app.route('/')
def index():
    # 🔒 se estiver logado, vai para a planilha
    if 'usuario' in session:
        return redirect(url_for('planilha_sig'))
    # 🔒 se não estiver logado, força login
    return redirect(url_for('login'))


# ================= RUN =================
# ================= RUN =================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

