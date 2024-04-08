from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect_with_retry("Extensão de arquivo não permitida. Permitido apenas arquivo em formato .xlsx")
        
        file = request.files['file']
        
        if file.filename == '':
            return redirect_with_retry("Extensão de arquivo não permitida. Permitido apenas arquivo em formato .xlsx")
        
        if file and allowed_file(file.filename):
            filename = file.filename
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            try:
                df = pd.read_excel(file_path)
                
                required_columns = [
                    'Tipo de Serviço', 'Operação', 'SOC', 'Hub', 'Rota', 'Período (Ano_Mês_Quinzena)',
                    '3PL Tracking Number / Número Etiqueta / Ordem (Shopee)',
                    '3PL Tracking Number (Enviado por API pela Transportadora)', 'Data da Prestação do Serviço',
                    'Nome Tomador do Serviço', 'CNPJ Tomador', 'Nome Transportadora (3PL)', 'CNPJ Emitente (3PL)',
                    'CEP Origem', 'Cidade Origem', 'UF Origem', 'CEP Entrega', 'Cidade Entrega', 'UF Entrega',
                    'Dados do Seller', 'Data da Entrega', 'Quantidade Volume', 'Peso Real', 'Peso Cubado',
                    'Peso Calculado/Cobrado', 'Km Rodado', 'Tipo do Veículo', 'Fator Agrupador (p/ Cobrança por Veiculo)',
                    'Placa', 'Valor nota Fiscal Mercadoria', 'Tarifa Aplicada', 'Frete/Tarifa Base (Peso/KM/Veículo)',
                    'Frete Calculado', 'ADV', 'GRIS', 'Aliquota ICMS/ISS', 'Base Calc ICMS/ISS', 'Valor ICMS/ISS',
                    'ICMS Subst', 'Outros Valores', 'Descontos', 'Valor Final à Receber', 'Data Emissão Cte/NF',
                    'Fatura', 'Número Cte', 'Número NF', 'Serie Cte/ Nf', 'Chave de Acesso Cte', 'Motivo Rejeição CTE',
                    'Prefeitura NFSe', 'Comentários', 'Payment Orders'
                ]
                
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    num_missing_columns = len(missing_columns)
                    error_message = f"Colunas ausentes: {num_missing_columns}\n\n"
                    error_message += "\n".join(missing_columns)
                    
                    return render_template('error.html', error_message=error_message, show_retry_button=True)
                
                # Se todas as colunas estiverem presentes, continue com o processamento
                return render_template('success.html', filename=filename)
            
            except Exception as e:
                return f'Ocorreu um erro ao processar o arquivo: {str(e)}'
        
        else:
            return redirect_with_retry("Extensão de arquivo não permitida. Permitido apenas arquivo em formato .xlsx")
    
    # Se o método for GET, renderize o template de upload
    return render_template('upload.html')


@app.route('/download_model', methods=['GET'])
def download_model():
    model_file = 'model_file.xlsx'
    model_path = os.path.join(app.root_path, model_file)
    
    if os.path.exists(model_path):
        return send_from_directory(app.root_path, model_file, as_attachment=True)
    else:
        return render_template('error.html', error_message='Arquivo modelo não encontrado.', show_retry_button=False), 404


def redirect_with_retry(error_message):
    return render_template('error.html', error_message=error_message, show_retry_button=True)


if __name__ == '__main__':
    app.run(debug=True)
# python app.py
