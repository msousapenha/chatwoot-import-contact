import os
import csv
import io
import requests
import openpyxl
from flask import Flask, request, redirect, url_for, session, flash, render_template_string

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY')

# Variáveis de Ambiente (Configuradas no Docker)
APP_USER = os.environ.get('APP_USER')
APP_PASS = os.environ.get('APP_PASS')
CHATWOOT_URL = os.environ.get('CHATWOOT_URL')
CHATWOOT_TOKEN = os.environ.get('CHATWOOT_TOKEN')

HTML_LOGIN = """
<!DOCTYPE html>
<html>
<head><title>Login - Importador</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="bg-light d-flex align-items-center justify-content-center" style="height: 100vh;">
    <div class="card p-4 shadow" style="width: 350px;">
        <h4 class="text-center mb-4">Importador Chatwoot</h4>
        {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
        <form method="POST">
            <div class="mb-3"><input type="text" name="username" class="form-control" placeholder="Usuário" required></div>
            <div class="mb-3"><input type="password" name="password" class="form-control" placeholder="Senha" required></div>
            <button type="submit" class="btn btn-primary w-100">Entrar</button>
        </form>
    </div>
</body>
</html>
"""

HTML_INDEX = """
<!DOCTYPE html>
<html>
<head><title>Importar Contatos</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="bg-light p-5">
    <div class="container max-w-md bg-white p-4 shadow rounded">
        <div class="d-flex justify-content-between align-items-center mb-4">
            <h2>Subir Planilha de Contatos</h2>
            <a href="{{ url_for('logout') }}" class="btn btn-sm btn-outline-danger">Sair</a>
        </div>
        <p class="text-muted">Envie um arquivo <b>.xlsx</b> (Excel) ou <b>.csv</b>. A planilha deve ter as colunas <b>Nome</b> e <b>Telefone</b>.</p>
        
        {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
        
        <form method="POST" enctype="multipart/form-data" class="mb-4">
            <input type="file" name="file" class="form-control mb-3" accept=".csv, .xlsx" required>
            <button type="submit" class="btn btn-success w-100">Processar e Importar</button>
        </form>

        {% if results %}
        <h5>Resultados:</h5>
        <ul class="list-group">
            {% for res in results %}
            <li class="list-group-item list-group-item-{{ 'success' if res.status == 'Sucesso' else 'danger' }}">
                <b>{{ res.nome }}</b>: {{ res.msg }}
            </li>
            {% endfor %}
        </ul>
        {% endif %}
    </div>
</body>
</html>
"""

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form['username'] == APP_USER and request.form['password'] == APP_PASS:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('Credenciais inválidas!', 'danger')
    return render_template_string(HTML_LOGIN)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('logged_in'): return redirect(url_for('login'))
    results = []

    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            flash('Nenhum arquivo enviado.', 'danger')
            return redirect(request.url)

        filename = file.filename.lower()
        if not (filename.endswith('.csv') or filename.endswith('.xlsx')):
            flash('Por favor, envie um arquivo .xlsx ou .csv válido', 'danger')
            return redirect(request.url)

        linhas_processadas = []

        # Lógica para ler XLSX
        if filename.endswith('.xlsx'):
            try:
                wb = openpyxl.load_workbook(file, data_only=True)
                sheet = wb.active
                rows = list(sheet.iter_rows(values_only=True))
                if len(rows) > 0:
                    headers_row = [str(h).strip().lower() if h else '' for h in rows[0]]
                    for row in rows[1:]:
                        if not any(row): continue # Pula linhas totalmente vazias
                        row_dict = dict(zip(headers_row, row))
                        linhas_processadas.append(row_dict)
            except Exception as e:
                flash(f'Erro ao ler o arquivo Excel: {str(e)}', 'danger')
                return redirect(request.url)
        
        # Lógica original para ler CSV mantida como fallback
        elif filename.endswith('.csv'):
            stream = io.StringIO(file.stream.read().decode("UTF8"), newline=None)
            reader = csv.DictReader(stream, delimiter=';') 
            for row in reader:
                row_lower = {k.strip().lower(): str(v).strip() for k, v in row.items() if k}
                linhas_processadas.append(row_lower)

        headers_api = {"api_access_token": CHATWOOT_TOKEN, "Content-Type": "application/json"}

        # Disparo para o Chatwoot
        for row_lower in linhas_processadas:
            nome = row_lower.get('nome') or row_lower.get('name')
            telefone = row_lower.get('telefone') or row_lower.get('phone_number')

            if not nome or not telefone:
                results.append({"nome": "Linha Inválida", "status": "Erro", "msg": "Colunas Nome ou Telefone ausentes."})
                continue
            
            # Limpa o telefone caso o Excel tenha formatado como número float (ex: 55869...0)
            telefone = str(telefone).replace('.0', '').strip()
            if not str(telefone).startswith('+'):
                telefone = f"+{telefone}"

            payload = {"name": str(nome).strip(), "phone_number": telefone}
            
            try:
                resp = requests.post(CHATWOOT_URL, json=payload, headers=headers_api)
                
                if resp.status_code == 200:
                    dados = resp.json().get('payload', {})
                    contact_id = dados.get('contact', {}).get('id') or dados.get('id')
                    
                    if contact_id:
                        url_etiqueta = f"{CHATWOOT_URL}/{contact_id}/labels"
                        requests.post(url_etiqueta, json={"labels": ["aceita_promocao"]}, headers=headers_api)
                        
                    results.append({"nome": nome, "status": "Sucesso", "msg": "Importado com a etiqueta 'aceita_promocao'!"})
                else:
                    results.append({"nome": nome, "status": "Erro", "msg": resp.json().get('message', resp.text)})
            except Exception as e:
                results.append({"nome": nome, "status": "Erro", "msg": str(e)})
        
        flash('Processamento concluído!', 'success')

    return render_template_string(HTML_INDEX, results=results)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)