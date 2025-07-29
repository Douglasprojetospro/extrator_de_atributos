from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from io import BytesIO
import os
from werkzeug.utils import secure_filename
import threading

# Configurações iniciais
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1GB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'sua_chave_secreta_aqui'  # Altere para uma chave segura

# Garante que a pasta de uploads exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Variável global para progresso
progresso = 0

def extrair_atributos(df_dados, df_config):
    global progresso
    resultado = df_dados.copy()
    regras = {}
    total_atributos = len(df_config['Atributo'].unique())
    
    for _, row in df_config.iterrows():
        atributo = row['Atributo']
        valor = row['Valor']
        padroes = [p.strip() for p in str(row['Padrões']).split(',') if p.strip()]
        
        if atributo not in regras:
            regras[atributo] = []
        regras[atributo].append({'valor': valor, 'padroes': padroes})
    
    for i, (atributo, regras_atributo) in enumerate(regras.items()):
        def aplicar_regras_row(desc):
            return aplicar_regras(desc, regras_atributo)
            
        resultado[atributo] = resultado['Descrição'].apply(aplicar_regras_row)
        
        progresso = int(((i + 1) / total_atributos) * 100)
    
    progresso = 100
    return resultado

def aplicar_regras(texto, regras):
    if pd.isna(texto) or not isinstance(texto, str):
        return None
        
    texto = texto.lower()
    for regra in regras:
        for padrao in regra['padroes']:
            if padrao.lower() in texto:
                return regra['valor']
    return None

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    global progresso
    
    try:
        progresso = 0
        
        if 'arquivo_dados' not in request.files or 'arquivo_config' not in request.files:
            return jsonify({"erro": "Envie ambos arquivos"}), 400
            
        arquivo_dados = request.files['arquivo_dados']
        arquivo_config = request.files['arquivo_config']
        
        if arquivo_dados.filename == '' or arquivo_config.filename == '':
            return jsonify({"erro": "Selecione arquivos válidos"}), 400
        
        dados_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo_dados.filename))
        config_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo_config.filename))
        
        arquivo_dados.save(dados_path)
        arquivo_config.save(config_path)
        
        def processar_thread():
            global progresso
            try:
                df_dados = pd.read_excel(dados_path)
                df_config = pd.read_excel(config_path)
                
                if 'Descrição' not in df_dados.columns:
                    return
                    
                required_cols = ['Atributo', 'Valor', 'Padrões']
                if not all(col in df_config.columns for col in required_cols):
                    return
                
                resultado = extrair_atributos(df_dados, df_config)
                resultado_path = os.path.join(app.config['UPLOAD_FOLDER'], 'resultado.xlsx')
                resultado.to_excel(resultado_path, index=False)
                
            except Exception as e:
                print(f"Erro no processamento: {str(e)}")
        
        threading.Thread(target=processar_thread).start()
        
        return jsonify({"mensagem": "Processamento iniciado"})
        
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route('/progresso')
def obter_progresso():
    return jsonify({"progresso": progresso})

@app.route('/download_resultado')
def download_resultado():
    resultado_path = os.path.join(app.config['UPLOAD_FOLDER'], 'resultado.xlsx')
    if os.path.exists(resultado_path):
        return send_file(
            resultado_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='resultados_processados.xlsx'
        )
    return jsonify({"erro": "Arquivo não disponível"}), 404

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
    app.run(host='0.0.0.0', port=5000)
