import os
import csv
import io
import requests
import urllib.parse
import openpyxl
import time
from datetime import datetime, timezone, timedelta
from flask import Flask, request, redirect, url_for, session, flash, render_template_string

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'a78s7hd1bnaxxmf21')

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
<head><title>Importar e Atualizar Contatos</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="bg-light p-5">
    <div class="container max-w-md bg-white p-4 shadow rounded">
        <div class="d-flex justify-content-between align-items-center mb-4">
            <h2>Subir Planilha de Contatos</h2>
            <a href="{{ url_for('logout') }}" class="btn btn-sm btn-outline-danger">Sair</a>
        </div>
        <p class="text-muted">Envie um arquivo <b>.xlsx</b> ou <b>.csv</b>. O sistema processará com limite de taxa (rate limit) de segurança.</p>
        
        {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
        
        <form method="POST" enctype="multipart/form-data" class="mb-4">
            <input type="file" name="file" class="form-control mb-3" accept=".csv, .xlsx" required>
            <button type="submit" class="btn btn-success w-100">Processar e Atualizar</button>
        </form>

        {% if results %}
        <h5>Resultados:</h5>
        <ul class="list-group">
            {% for res in results %}
            <li class="list-group-item list-group-item-{{ 'success' if res.status == 'Sucesso' else 'warning' if res.status == 'Atualizado' else 'danger' }}">
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
            return redirect(request.url)

        filename = file.filename.lower()
        linhas_processadas = []

        if filename.endswith('.xlsx'):
            try:
                wb = openpyxl.load_workbook(file, data_only=True)
                sheet = wb.active
                rows = list(sheet.iter_rows(values_only=True))
                
                header_idx = -1
                for i, row in enumerate(rows):
                    if not any(row): continue
                    row_str = [str(c).strip().lower() for c in row if c is not None]
                    if any(x in row_str for x in ['nome', 'name', 'paciente', 'contato']) and any(x in row_str for x in ['telefone', 'celular', 'phone', 'whatsapp']):
                        header_idx = i
                        break
                
                if header_idx == -1:
                    flash('Erro: Não encontrei as colunas de Nome e Telefone na planilha.', 'danger')
                    return redirect(request.url)
                    
                headers_row = [str(h).strip().lower() if h else f'vazio_{i}' for i, h in enumerate(rows[header_idx])]
                
                for row in rows[header_idx+1:]:
                    if not any(row): continue
                    row_dict = dict(zip(headers_row, row))
                    linhas_processadas.append(row_dict)
            except Exception as e:
                flash(f'Erro ao ler Excel: {str(e)}', 'danger')
                return redirect(request.url)
                
        elif filename.endswith('.csv'):
            content = file.stream.read().decode("utf-8-sig")
            stream = io.StringIO(content, newline=None)
            reader = csv.DictReader(stream, delimiter=';')
            if not reader.fieldnames or len(reader.fieldnames) < 2:
                stream.seek(0)
                reader = csv.DictReader(stream, delimiter=',')
                
            for row in reader:
                row_lower = {k.strip().lower(): str(v).strip() for k, v in row.items() if k}
                linhas_processadas.append(row_lower)

        headers_api = {"api_access_token": CHATWOOT_TOKEN, "Content-Type": "application/json"}
        
        # Inicializa a Sessão HTTP (Connection Pooling)
        http_session = requests.Session()
        http_session.headers.update(headers_api)

        fuso_brasil = timezone(timedelta(hours=-3))
        data_atual = datetime.now(fuso_brasil).strftime("%d-%m-%Y")
        nome_da_etiqueta = f"novos_contatos_{data_atual}"
        
        url_criar_etiqueta = CHATWOOT_URL.replace('/contacts', '/labels')
        
        payload_nova_etiqueta = {
            "title": nome_da_etiqueta,
            "description": f"Contatos importados na data {data_atual}",
            "color": "#28a745",
            "show_on_sidebar": True
        }
        
        try:
            http_session.post(url_criar_etiqueta, json=payload_nova_etiqueta)
        except Exception as e:
            print(f"Erro ignorado ao criar etiqueta: {e}")

        for row_lower in linhas_processadas:
            nome = row_lower.get('nome') or row_lower.get('paciente') or row_lower.get('name')
            telefone = row_lower.get('telefone') or row_lower.get('celular') or row_lower.get('phone_number')
            vendedora = row_lower.get('vendedora') or row_lower.get('vendedor')

            if not nome or not telefone or str(telefone).lower() == 'none':
                continue
            
            telefone = str(telefone).replace('.0', '').replace(' ', '').replace('-', '').replace('(', '').replace(')', '').strip()
            if not telefone.startswith('+'):
                telefone = f"+{telefone}"

            contact_id = None
            is_update = False

            try:
                phone_encoded = urllib.parse.quote(telefone)
                search_url = f"{CHATWOOT_URL}/search?q={phone_encoded}"
                
                resp_search = http_session.get(search_url)
                if resp_search.status_code == 200:
                    search_data = resp_search.json().get('payload', [])
                    if len(search_data) > 0:
                        contact_id = search_data[0].get('id')
                        is_update = True
            except Exception as e:
                print(f"Erro na busca: {e}")

            if not contact_id:
                payload_create = {
                    "name": str(nome).strip(), 
                    "phone_number": telefone
                }
                
                try:
                    resp_create = http_session.post(CHATWOOT_URL, json=payload_create)
                    if resp_create.status_code == 200:
                        dados = resp_create.json().get('payload', {})
                        contact_id = dados.get('contact', {}).get('id') or dados.get('id')
                    else:
                        results.append({"nome": nome, "status": "Erro", "msg": f"Erro ao criar: {resp_create.text}"})
                        continue
                except Exception as e:
                    results.append({"nome": nome, "status": "Erro", "msg": str(e)})
                    continue

            if contact_id:
                if vendedora and str(vendedora).strip() and str(vendedora).lower() != 'none':
                    url_update = f"{CHATWOOT_URL}/{contact_id}"
                    payload_update = {
                        "custom_attributes": {
                            "vendedora": str(vendedora).strip()
                        }
                    }
                    http_session.put(url_update, json=payload_update)
                
                url_etiqueta_contato = f"{CHATWOOT_URL}/{contact_id}/labels"
                http_session.post(url_etiqueta_contato, json={"labels": ["aceita_promocao", nome_da_etiqueta]})
                
                status_texto = "Atualizado" if is_update else "Sucesso"
                msg_texto = "Contato existente atualizado com as novas etiquetas!" if is_update else "Novo contato criado com sucesso!"
                results.append({"nome": nome, "status": status_texto, "msg": msg_texto})
            
            # Throttling: Pausa de 0.5 segundos para aliviar o Traefik/Chatwoot e prevenir bloqueios
            time.sleep(0.5)
        
        flash('Processamento concluído!', 'success')

    return render_template_string(HTML_INDEX, results=results)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)