from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from io import BytesIO
import os
from werkzeug.utils import secure_filename
import threading
import atexit
import shutil

# Configurações iniciais
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1GB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = os.urandom(24)  # Chave secreta mais segura
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Garante que a pasta de uploads exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Variáveis globais com thread-safe
progresso = 0
progresso_lock = threading.Lock()
processo_em_andamento = False

# Limpeza ao sair
def cleanup():
    try:
        shutil.rmtree(app.config['UPLOAD_FOLDER'])
    except:
        pass

atexit.register(cleanup)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def extrair_atributos(df_dados, df_config):
    global progresso
    resultado = df_dados.copy()
    regras = {}
    total_atributos = len(df_config['Atributo'].unique())
    
    # Pré-processamento das regras
    for _, row in df_config.iterrows():
        atributo = row['Atributo']
        valor = str(row['Valor'])
        padroes = [p.strip().lower() for p in str(row['Padrões']).split(',') if p.strip()]
        
        if atributo not in regras:
            regras[atributo] = []
        regras[atributo].append({'valor': valor, 'padroes': padroes})
    
    # Aplicação das regras com progresso
    for i, (atributo, regras_atributo) in enumerate(regras.items()):
        resultado[atributo] = resultado['Descrição'].apply(
            lambda desc: aplicar_regras(desc, regras_atributo)
        
        with progresso_lock:
            progresso = int(((i + 1) / total_atributos) * 100)
    
    with progresso_lock:
        progresso = 100
    
    return resultado

def aplicar_regras(texto, regras):
    if pd.isna(texto) or not isinstance(texto, str):
        return None
        
    texto = texto.lower()
    for regra in regras:
        for padrao in regra['padroes']:
            if padrao in texto:
                return regra['valor']
    return None

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    global progresso, processo_em_andamento
    
    try:
        with progresso_lock:
            if processo_em_andamento:
                return jsonify({"erro": "Já existe um processamento em andamento"}), 429
            processo_em_andamento = True
            progresso = 0
        
        if 'arquivo_dados' not in request.files or 'arquivo_config' not in request.files:
            return jsonify({"erro": "Envie ambos arquivos"}), 400
            
        arquivo_dados = request.files['arquivo_dados']
        arquivo_config = request.files['arquivo_config']
        
        if arquivo_dados.filename == '' or arquivo_config.filename == '':
            return jsonify({"erro": "Selecione arquivos válidos"}), 400
        
        if not (allowed_file(arquivo_dados.filename) and allowed_file(arquivo_config.filename)):
            return jsonify({"erro": "Apenas arquivos Excel são permitidos"}), 400
        
        # Criar subpasta para esta sessão
        session_id = secure_filename(str(threading.get_ident()))
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)
        
        dados_path = os.path.join(session_folder, secure_filename(arquivo_dados.filename))
        config_path = os.path.join(session_folder, secure_filename(arquivo_config.filename))
        resultado_path = os.path.join(session_folder, 'resultado.xlsx')
        
        arquivo_dados.save(dados_path)
        arquivo_config.save(config_path)
        
        def processar_thread():
            global progresso, processo_em_andamento
            try:
                # Ler arquivos com tratamento de erro
                try:
                    df_dados = pd.read_excel(dados_path)
                    df_config = pd.read_excel(config_path)
                except Exception as e:
                    with progresso_lock:
                        progresso = -1  # Indica erro
                    return
                
                # Verificar colunas necessárias
                if 'Descrição' not in df_dados.columns:
                    with progresso_lock:
                        progresso = -2  # Erro de coluna faltante
                    return
                    
                required_cols = ['Atributo', 'Valor', 'Padrões']
                if not all(col in df_config.columns for col in required_cols):
                    with progresso_lock:
                        progresso = -3  # Erro de configuração
                    return
                
                # Processamento principal
                resultado = extrair_atributos(df_dados, df_config)
                
                # Salvar resultado
                resultado.to_excel(resultado_path, index=False)
                
            except Exception as e:
                print(f"Erro no processamento: {str(e)}")
                with progresso_lock:
                    progresso = -4  # Erro genérico
            finally:
                with progresso_lock:
                    processo_em_andamento = False
        
        threading.Thread(target=processar_thread).start()
        
        return jsonify({
            "mensagem": "Processamento iniciado",
            "session_id": session_id
        })
        
    except Exception as e:
        with progresso_lock:
            processo_em_andamento = False
        return jsonify({"erro": str(e)}), 500

@app.route('/progresso')
def obter_progresso():
    with progresso_lock:
        return jsonify({
            "progresso": progresso,
            "status": "processando" if progresso < 100 and progresso >= 0 else "erro",
            "mensagem_erro": get_error_message(progresso)
        })

def get_error_message(error_code):
    errors = {
        -1: "Erro ao ler arquivos Excel",
        -2: "Coluna 'Descrição' não encontrada no arquivo de dados",
        -3: "Colunas obrigatórias não encontradas no arquivo de configuração",
        -4: "Erro durante o processamento"
    }
    return errors.get(error_code, "")

@app.route('/download_resultado/<session_id>')
def download_resultado(session_id):
    session_id = secure_filename(session_id)
    resultado_path = os.path.join(app.config['UPLOAD_FOLDER'], session_id, 'resultado.xlsx')
    
    if not os.path.exists(resultado_path):
        return jsonify({"erro": "Arquivo não disponível"}), 404
    
    try:
        return send_file(
            resultado_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='resultados_processados.xlsx'
        )
    finally:
        # Limpar arquivos após download
        shutil.rmtree(os.path.join(app.config['UPLOAD_FOLDER'], session_id), ignore_errors=True)

@app.route('/download_modelo_produtos')
def download_modelo_produtos():
    dados = {
        'ID': [1, 2, 3, 4, 5],
        'Descrição': [
            "Liquidificador Mondial 110V 500W cor branca",
            "Ventilador Arno 220V com 3 velocidades",
            "Fogão Consul 4 bocas cor inox",
            "Micro-ondas Panasonic 20L 110V",
            "Geladeira Brastemp Frost Free 375L"
        ],
        'Categoria': ["Eletroportátil", "Eletroportátil", "Eletrodoméstico", "Eletrodoméstico", "Eletrodoméstico"]
    }
    df = pd.DataFrame(dados)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='modelo_produtos.xlsx'
    )

@app.route('/download_modelo_config')
def download_modelo_config():
    dados = {
        'Atributo': ['Voltagem', 'Voltagem', 'Cor', 'Cor', 'Tamanho', 'Potência', 'Marca'],
        'Valor': ['110V', '220V', 'Branco', 'Preto', 'Grande', '500W', 'Mondial'],
        'Padrões': [
            '110, 110v, 110 volts',
            '220, 220v, 220 volts',
            'branco, white, branca',
            'preto, black, pretinha',
            'grande, large, xl, gg',
            '500, 500w, 500 watts',
            'mondial, arno, consul'
        ]
    }
    df = pd.DataFrame(dados)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='modelo_configuracao.xlsx'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, threaded=True)
