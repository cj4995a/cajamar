from flask import Flask, render_template_string, request, redirect
from datetime import datetime

app = Flask(__name__)

# =========================================================
# UNIDADES
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

# =========================================================
# CONFIG
# =========================================================

CONFIG_PADRAO = {
    key: {
        "HORARIOS": [
            {
                "hora": "08:00",
                "dias": ["0", "1", "2", "3", "4"]
            }
        ],
        "STATUS": "Parado",
        "ULTIMA_EXECUCAO": "-"
    }
    for key in UNIDADES_CONFIG
}

CONFIG_GLOBAL = CONFIG_PADRAO.copy()

def carregar_config():
    global CONFIG_GLOBAL
    return CONFIG_GLOBAL

def salvar_config(config):
    global CONFIG_GLOBAL
    CONFIG_GLOBAL = config

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
}

*{
    margin:0;
    padding:0;
    box-sizing:border-box;
}

body{
    background:var(--bg);
    color:var(--text);
    font-family:Arial;
    padding:30px;
}

h1{
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
    border-radius:20px;
    padding:20px;
    border:1px solid var(--border);
}

h2{
    margin-bottom:15px;
}

.info{
    color:var(--sub);
    margin-bottom:18px;
    line-height:1.8;
}

.horario{
    border:1px solid var(--border);
    border-radius:14px;
    padding:12px;
    margin-bottom:12px;
}

input[type="time"]{
    width:100%;
    padding:12px;
    border:none;
    border-radius:10px;
    background:#1e293b;
    color:white;
    margin-bottom:10px;
}

.dias{
    display:flex;
    flex-wrap:wrap;
    gap:8px;
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
    font-weight:bold;
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

<div class="info">

<div>
<b>Status:</b>
{{ dados.STATUS }}
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

    config = carregar_config()

    return render_template_string(
        HTML,
        config=config,
        unidades_config=UNIDADES_CONFIG,
        dias_semana=DIAS_SEMANA
    )

@app.route("/salvar_horarios/<unidade>", methods=["POST"])
def salvar_horarios(unidade):

    config = carregar_config()

    horarios = []

    horas = request.form.getlist("hora")

    for i, hora in enumerate(horas):

        dias = request.form.getlist(f"dias_{i}")

        horarios.append({
            "hora": hora,
            "dias": dias
        })

    config[unidade]["HORARIOS"] = horarios

    config[unidade]["ULTIMA_EXECUCAO"] = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    salvar_config(config)

    return redirect("/")

# =========================================================
# VERCEL
# =========================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
