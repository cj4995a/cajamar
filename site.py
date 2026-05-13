from flask import Flask, render_template_string, request, redirect
import os
import json
import subprocess
import threading
import time
from datetime import datetime
import psutil
import logging
import sys

app = Flask(__name__)

# =========================================================
# LOG FLASK
# =========================================================

log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)

# =========================================================
# PASTAS
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
LOG_DIR = os.path.join(BASE_DIR, "logs")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

CONFIG_FILE = os.path.join(DATA_DIR, "config.json")

# =========================================================
# SCRIPTS
# =========================================================

SCRIPTS = {
    "DEMANDA_OP_CAJAMAR": os.path.join(BASE_DIR, "demanda_op_cajamar.py"),
    "PROD_REFRIO": os.path.join(BASE_DIR, "prod_refrio.py"),
    "SAMS_BRASIL": os.path.join(BASE_DIR, "sams_brasil.py"),
}

# =========================================================
# VISUAL
# =========================================================

UNIDADES_CONFIG = {

    "DEMANDA_OP_CAJAMAR": {
        "icone": "📦",
        "label": "Demanda OP Cajamar",
    },

    "PROD_REFRIO": {
        "icone": "❄️",
        "label": "Prod Refrio",
    },

    "SAMS_BRASIL": {
        "icone": "🛒",
        "label": "Sams Brasil",
    },
}

DIAS_SEMANA = [
    ('0', 'Seg'),
    ('1', 'Ter'),
    ('2', 'Qua'),
    ('3', 'Qui'),
    ('4', 'Sex'),
    ('5', 'Sáb'),
    ('6', 'Dom')
]

PROCESSOS = {}
ULTIMA_EXECUCAO = {}

# =========================================================
# CONFIG
# =========================================================

def carregar_configuracoes():

    if not os.path.exists(CONFIG_FILE):

        config = {
            key: {
                "HORARIOS": [
                    {
                        "hora": "08:00",
                        "dias": ["0", "1", "2", "3", "4"]
                    }
                ],
                "STATUS": "Parado",
                "PID": None,
                "ULTIMA_EXECUCAO": "-",
                "ONLINE": False
            }
            for key in UNIDADES_CONFIG
        }

        salvar_configuracoes(config)

        return config

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_configuracoes(config):

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

# =========================================================
# LOGS
# =========================================================

def escrever_log(unidade, mensagem):

    arquivo = os.path.join(LOG_DIR, f"{unidade}.log")

    linha = (
        f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] "
        f"{mensagem}\n"
    )

    with open(arquivo, "a", encoding="utf-8") as f:
        f.write(linha)

def ler_logs(unidade):

    arquivo = os.path.join(LOG_DIR, f"{unidade}.log")

    if not os.path.exists(arquivo):
        return "Sem logs"

    with open(arquivo, "r", encoding="utf-8") as f:
        linhas = f.readlines()

    if not linhas:
        return "Sem logs"

    return linhas[-1]
# =========================================================
# EXECUTAR SCRIPT
# =========================================================

def executar_script(unidade):

    unidade = unidade.upper()

    script = SCRIPTS.get(unidade)

    if not script:
        escrever_log(unidade, "Script não encontrado")
        return

    if not os.path.exists(script):
        escrever_log(unidade, f"Arquivo não encontrado: {script}")
        return

    processo_existente = PROCESSOS.get(unidade)

    if processo_existente:

        try:

            if processo_existente.poll() is None:

                escrever_log(
                    unidade,
                    "Já está executando"
                )

                return

        except:
            pass

    try:

        escrever_log(
            unidade,
            "Iniciando automação..."
        )

        python_exe = sys.executable

        processo = subprocess.Popen(

            [
                python_exe,
                script
            ],

            cwd=BASE_DIR,

            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,

            creationflags=subprocess.CREATE_NO_WINDOW
        )

        PROCESSOS[unidade] = processo

        config = carregar_configuracoes()

        config[unidade]["STATUS"] = "Executando"
        config[unidade]["ONLINE"] = True
        config[unidade]["PID"] = processo.pid

        config[unidade]["ULTIMA_EXECUCAO"] = datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        salvar_configuracoes(config)

        escrever_log(
            unidade,
            f"PID iniciado {processo.pid}"
        )

    except Exception as e:

        escrever_log(
            unidade,
            f"ERRO: {e}"
        )

# =========================================================
# PARAR SCRIPT
# =========================================================

def parar_script(unidade):

    unidade = unidade.upper()

    processo = PROCESSOS.get(unidade)

    try:

        if processo and processo.poll() is None:

            parent = psutil.Process(processo.pid)

            for child in parent.children(recursive=True):
                child.kill()

            parent.kill()

            escrever_log(
                unidade,
                "Processo finalizado"
            )

    except Exception as e:

        escrever_log(
            unidade,
            f"Erro STOP: {e}"
        )

    config = carregar_configuracoes()

    config[unidade]["STATUS"] = "Parado"
    config[unidade]["ONLINE"] = False
    config[unidade]["PID"] = None

    salvar_configuracoes(config)

# =========================================================
# SCHEDULER
# =========================================================

def scheduler():

    while True:

        agora = datetime.now()

        hora_atual = agora.strftime("%H:%M")

        dia_semana = str(agora.weekday())

        config = carregar_configuracoes()

        for unidade, dados in config.items():

            horarios = dados.get("HORARIOS", [])

            for agenda in horarios:

                hora = agenda.get("hora", "").strip()

                dias = agenda.get("dias", [])

                chave_execucao = f"{unidade}_{hora}_{agora.strftime('%Y%m%d')}"

                if (
                    hora == hora_atual
                    and dia_semana in dias
                ):

                    if ULTIMA_EXECUCAO.get(chave_execucao):
                        continue

                    processo = PROCESSOS.get(unidade)

                    if processo:

                        try:
                            if processo.poll() is None:
                                continue
                        except:
                            pass

                    escrever_log(
                        unidade,
                        f"Horário automático disparado: {hora}"
                    )

                    threading.Thread(
                        target=executar_script,
                        args=(unidade,),
                        daemon=True
                    ).start()

                    ULTIMA_EXECUCAO[chave_execucao] = True

        time.sleep(1)

# =========================================================
# HTML
# =========================================================

HTML = """

<!DOCTYPE html>

<html lang="pt-br">

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width, initial-scale=1.0">

<title>Painel PCP</title>

<style>

:root{
    --bg:#0f172a;
    --card:#111827;
    --border:rgba(255,255,255,0.08);
    --text:#f8fafc;
    --sub:#94a3b8;
    --green:#22c55e;
    --blue:#3b82f6;
    --red:#ef4444;
}

*{
    margin:0;
    padding:0;
    box-sizing:border-box;
}

body{
    background:var(--bg);
    color:var(--text);
    font-family:Inter,Segoe UI,Arial,sans-serif;
    padding:30px;
}

h1{
    font-size:32px;
    margin-bottom:25px;
}

.container{
    display:grid;
    grid-template-columns:
    repeat(auto-fit,minmax(320px,1fr));
    gap:20px;
}

.card{
    background:var(--card);
    border:1px solid var(--border);
    border-radius:20px;
    padding:18px;
    width:100%;
}

.card h2{
    margin-bottom:12px;
    font-size:20px;
}

.status{
    display:inline-block;
    padding:7px 14px;
    border-radius:999px;
    font-size:12px;
    font-weight:700;
    margin-bottom:16px;
}

.online{
    background:rgba(34,197,94,0.15);
    color:#22c55e;
}

.offline{
    background:rgba(239,68,68,0.15);
    color:#ef4444;
}

.info{
    display:flex;
    flex-direction:column;
    gap:8px;
    margin-bottom:16px;
    color:var(--sub);
    font-size:14px;
}

.horario{
    border:1px solid var(--border);
    border-radius:16px;
    padding:14px;
    margin-bottom:12px;
}

input[type="time"]{
    width:100%;
    padding:12px;
    border-radius:12px;
    border:none;
    background:#1e293b;
    color:white;
    margin-bottom:10px;
}

.dias{
    display:flex;
    flex-wrap:wrap;
    gap:10px;
}

.dias label{
    font-size:13px;
    color:var(--sub);
}

button{
    width:100%;
    padding:14px;
    border:none;
    border-radius:14px;
    cursor:pointer;
    font-weight:700;
    margin-top:10px;
}

.add{
    background:#334155;
    color:white;
}

.save{
    background:#22c55e;
    color:white;
}

.run{
    background:#3b82f6;
    color:white;
}

.stop{
    background:#ef4444;
    color:white;
}

.logs{
    margin-top:14px;
    background:#020617;
    border-radius:14px;
    padding:12px;
    height:180px;

    overflow-y:auto;
    overflow-x:hidden;

    white-space:pre-wrap;
    word-break:break-word;
    overflow-wrap:anywhere;

    font-size:12px;
    line-height:1.5;

    color:#e2e8f0;
}

.logs::-webkit-scrollbar{
    width:6px;
}

.logs::-webkit-scrollbar-thumb{
    background:#334155;
    border-radius:999px;
}

</style>

</head>

<body>

<h1>⚙️ Painel PCP</h1>

<div class="container">

{% for unidade, dados in config.items() %}

<div class="card">

<h2>
{{ unidades_config[unidade].icone }}
{{ unidades_config[unidade].label }}
</h2>

{% if dados.ONLINE %}
<div class="status online">ONLINE</div>
{% else %}
<div class="status offline">OFFLINE</div>
{% endif %}

<div class="info">

<div>
<b>Status:</b>
{{ dados.STATUS }}
</div>

<div>
<b>PID:</b>
{{ dados.PID }}
</div>

<div>
<b>Última execução:</b>
{{ dados.ULTIMA_EXECUCAO }}
</div>

</div>

<form
action="/salvar_horarios/{{ unidade }}"
method="post"
>

<div id="container-{{ unidade }}">

{% for horario in dados.HORARIOS %}

{% set idx = loop.index0 %}

<div class="horario">

<input
type="time"
name="hora"
value="{{ horario.hora }}"
required
>

<div class="dias">

{% for val, nome in dias_semana %}

<label>

<input
type="checkbox"
name="dias_{{ idx }}"
value="{{ val }}"
{{ 'checked' if val in horario.dias }}
>

{{ nome }}

</label>

{% endfor %}

</div>

</div>

{% endfor %}

</div>

<button
type="button"
class="add"
onclick="adicionarHorario('{{ unidade }}')"
>

➕ Adicionar Horário

</button>

<button class="save">

💾 Salvar Horários

</button>

</form>

{% if dados.ONLINE %}

<form
action="/parar/{{ unidade }}"
method="post"
>

<button class="stop">

⛔ Parar

</button>

</form>

{% else %}

<form
action="/executar/{{ unidade }}"
method="post"
>

<button class="run">

▶️ Executar Agora

</button>

</form>

{% endif %}

<div class="logs">

{{ logs[unidade] }}

</div>

</div>

{% endfor %}

</div>

<script>

function adicionarHorario(unidade){

    let container =
    document.getElementById(
        `container-${unidade}`
    )

    let total = container.children.length

    let html = `

    <div class="horario">

    <input type="time" name="hora" required>

    <div class="dias">

    <label><input type="checkbox" name="dias_${total}" value="0">Seg</label>

    <label><input type="checkbox" name="dias_${total}" value="1">Ter</label>

    <label><input type="checkbox" name="dias_${total}" value="2">Qua</label>

    <label><input type="checkbox" name="dias_${total}" value="3">Qui</label>

    <label><input type="checkbox" name="dias_${total}" value="4">Sex</label>

    <label><input type="checkbox" name="dias_${total}" value="5">Sáb</label>

    <label><input type="checkbox" name="dias_${total}" value="6">Dom</label>

    </div>

    </div>

    `

    container.insertAdjacentHTML(
        'beforeend',
        html
    )
}

</script>

</body>

</html>

"""

# =========================================================
# ROTAS
# =========================================================

@app.route("/")
def index():

    config = carregar_configuracoes()

    for unidade in config:

        processo = PROCESSOS.get(unidade)

        if processo:

            try:

                if processo.poll() is None:

                    config[unidade]["ONLINE"] = True
                    config[unidade]["STATUS"] = "Executando"

                else:

                    config[unidade]["ONLINE"] = False
                    config[unidade]["STATUS"] = "Parado"
                    config[unidade]["PID"] = None

            except:
                pass

    salvar_configuracoes(config)

    logs = {}

    for unidade in config:
        logs[unidade] = ler_logs(unidade)

    return render_template_string(
        HTML,
        config=config,
        unidades_config=UNIDADES_CONFIG,
        dias_semana=DIAS_SEMANA,
        logs=logs
    )

@app.route("/executar/<unidade>", methods=["POST"])
def executar(unidade):

    threading.Thread(
        target=executar_script,
        args=(unidade,),
        daemon=True
    ).start()

    return redirect("/")

@app.route("/parar/<unidade>", methods=["POST"])
def parar(unidade):

    parar_script(unidade)

    return redirect("/")

@app.route("/salvar_horarios/<unidade>", methods=["POST"])
def salvar_horarios(unidade):

    config = carregar_configuracoes()

    horarios = []

    horas = request.form.getlist("hora")

    for i, hora in enumerate(horas):

        dias = request.form.getlist(f"dias_{i}")

        horarios.append({
            "hora": hora,
            "dias": dias
        })

    config[unidade]["HORARIOS"] = horarios

    salvar_configuracoes(config)

    escrever_log(
        unidade,
        "Horários atualizados"
    )

    return redirect("/")

# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":

    threading.Thread(
        target=scheduler,
        daemon=True
    ).start()

    print("Painel iniciado")
    print("http://localhost:8000")

    app.run(
        host="0.0.0.0",
        port=8000,
        debug=False
    )