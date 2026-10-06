import os
import pandas as pd
from datetime import date, timedelta, datetime
from dotenv import load_dotenv
import pyodbc 
from jinja2 import Template
import pdfkit
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 1. Carregar variáveis de ambiente
load_dotenv()
DB_SERVER = os.getenv("DB_SERVER")
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASS = os.getenv("DB_PASS")
DB_DRIVER = os.getenv("DB_DRIVER")

EMAIL_HOST = os.getenv("EMAIL_HOST")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", 587))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
EMAIL_SUPORTE = "ti@skyglasscanela.com.br"

# --- FUNÇÃO PARA ENVIAR E-MAIL ---
def enviar_email_com_anexo(destinatario, assunto, corpo, caminho_anexo=None):
    if not destinatario or str(destinatario).strip() == "" or str(destinatario).lower() == "nan":
        print("   -> E-mail ignorado: Destinatário inválido.")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = str(destinatario).strip()
    msg['Subject'] = assunto

    msg.attach(MIMEText(corpo, 'html'))

    # Anexando arquivo se houver
    if caminho_anexo and os.path.exists(caminho_anexo):
        try:
            with open(caminho_anexo, "rb") as f:
                anexo = MIMEApplication(f.read(), _subtype="pdf" if caminho_anexo.endswith('.pdf') else "plain")
                anexo.add_header('Content-Disposition', 'attachment', filename=os.path.basename(caminho_anexo))
                msg.attach(anexo)
        except Exception as e:
            print(f"   -> Erro ao ler o anexo: {e}")
            return False

    try:
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"   -> Erro ao enviar e-mail para {destinatario}: {e}")
        return False
# ---------------------------------

# 3. Definição Automática do Período (Mês Anterior)
hoje = date.today()
primeiro_dia_atual = hoje.replace(day=1)
ultimo_dia_mes_anterior = primeiro_dia_atual - timedelta(days=1)
primeiro_dia_mes_anterior = ultimo_dia_mes_anterior.replace(day=1)

str_primeiro_dia = primeiro_dia_mes_anterior.strftime('%Y-%m-%d')
str_corte_mes_atual = primeiro_dia_atual.strftime('%Y-%m-%d') 
str_ultimo_dia_mes_anterior = ultimo_dia_mes_anterior.strftime('%Y-%m-%d')

data_inicial_fmt = primeiro_dia_mes_anterior.strftime('%d/%m/%Y')
data_final_fmt = ultimo_dia_mes_anterior.strftime('%d/%m/%Y')
mes_atual_fmt = primeiro_dia_atual.strftime('%m')

meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
nome_mes_anterior = meses[primeiro_dia_mes_anterior.month - 1]
mes_referencia = f"{nome_mes_anterior}/{primeiro_dia_mes_anterior.year}"
data_emissao = hoje.strftime('%d/%m/%Y')

# 2. Configurações de Diretórios (Organizado por Ano.Mês, ex: relatorios/2026.10)
pasta_ano_mes = primeiro_dia_mes_anterior.strftime('%Y.%m')
PASTA_RELATORIOS = os.path.join('relatorios', pasta_ano_mes)
if not os.path.exists(PASTA_RELATORIOS):
    os.makedirs(PASTA_RELATORIOS)

caminho_logo = os.path.abspath('img/logo.png')
with open(caminho_logo, "rb") as image_file:
    logo_codificada = base64.b64encode(image_file.read()).decode('utf-8')
logo_base64 = f"data:image/png;base64,{logo_codificada}"

print(f"Buscando comissões gerais de {str_primeiro_dia} até {str_ultimo_dia_mes_anterior}...")

# 4. Consulta SQL Geral
query = f"""
SELECT 
    PEDPLUS.FK_GUIA AS Cod_Parcei,
    GUIA.DS_FANTASIA AS Parceiro,
    GUIA.nr_cnpj AS Cnpj,           
    GUIA.ds_email AS Email,         
    PED.DT_PEDIDO AS Dt_Venda,
    PED.NR_PEDIDO AS Nro_Venda,
    PED.VL_TOTAL AS Vl_Total,
    MOV.DS_MODELO AS Modelo,
    (ABS(MOV.QT_MOVIMENTO) * (CASE WHEN MOV.TG_ES = 'S' THEN 1 ELSE -1 END)) AS Qtd,
    (ABS(MOV.VL_PREUNI) * (CASE WHEN MOV.TG_ES = 'S' THEN 1 ELSE -1 END)) AS Vl_Unitario,
    MOV.VL_PRETOT AS Preco_Total,
    COALESCE(GUIA.VL_COMISSAO, 0) AS Comissao_Perc,
    COALESCE(ROUND(GUIA.VL_COMISSAO * MOV.VL_PRETOT / 100, 2), 0) AS Valor_Comissao,
    CASE 
        WHEN PED.DS_COMPLEMENTO IS NULL OR LTRIM(RTRIM(PED.DS_COMPLEMENTO)) = '' 
        THEN 'NÃO IDENTIFICADO' 
        ELSE PED.DS_COMPLEMENTO 
    END AS Cliente
FROM VD_PEDIDOS PED
INNER JOIN VD_PEDIDOSPLUS PEDPLUS ON PED.PK_ID = PEDPLUS.PK_ID
INNER JOIN TB_CADUNICO GUIA ON GUIA.PK_ID = PEDPLUS.FK_GUIA
INNER JOIN ES_MOVIMENTOS MOV ON MOV.FK_ORIGEM = PED.PK_ID AND MOV.TG_ORIGEM = 'V'
WHERE PEDPLUS.FK_GUIA > 0
  AND PED.DT_PEDIDO >= '{str_primeiro_dia}'
  AND PED.DT_PEDIDO < '{str_corte_mes_atual}'
  AND COALESCE(PED.TG_INATIVO, 0) = 0
  AND COALESCE(MOV.TG_INATIVO, 0) = 0
ORDER BY PEDPLUS.FK_GUIA, PED.DT_PEDIDO, PED.NR_PEDIDO
"""

# Template HTML do Relatório PDF
html_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; font-size: 11px; margin: 20px; }
        .header { text-align: center; margin-bottom: 25px; position: relative; }
        .logo { position: absolute; left: 0; top: 0; height: 65px; }
        h2 { margin: 0 0 5px 0; font-size: 18px; font-weight: normal; }
        p { margin: 3px 0; font-size: 13px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 0; }
        th { background-color: #1f4e79; color: white; border: 1px solid #000; padding: 6px; text-align: center; font-size: 10px; }
        td { border: 1px solid #000; padding: 5px; text-align: center; font-size: 10px; }
        .text-left { text-align: left; }
        .footer-row td { background-color: #fff2cc; font-weight: bold; border-top: 2px solid #000; }
    </style>
</head>
<body>
    <div class="header">
        <img src="{{ logo_base64 }}" class="logo" />
        <h2>BOLETIM DE COMISSÃO</h2>
        <p>{{ nome_parceiro }} (Cód: {{ codigo_parceiro }})</p>
        <p>CNPJ: {{ cnpj_parceiro }} // E-mail: {{ email_parceiro }}</p> 
        <p>Referência: {{ mes_referencia }} // Emitido: {{ data_emissao }}</p>
    </div>
    <table>
        <thead>
            <tr>
                <th>Cód.Parcei.</th>
                <th>Parceiro</th>
                <th>Dt.Venda</th>
                <th>Nro.Venda</th>
                <th>Vl.Total</th>
                <th>Modelo</th>
                <th>Qtd</th>
                <th>Vl. Unitário</th>
                <th>Preço Total</th>
                <th>Comissão</th>
                <th>Valor comissão</th>
                <th>Cliente</th>
            </tr>
        </thead>
        <tbody>
            {% for index, row in itens.iterrows() %}
            <tr>
                <td>{{ row['Cod_Parcei'] }}</td>
                <td class="text-left">{{ row['Parceiro'] }}</td>
                <td>{{ row['Dt_Venda'] }}</td>
                <td>{{ row['Nro_Venda'] }}</td>
                <td>{{ row['Vl_Total'] }}</td>
                <td class="text-left">{{ row['Modelo'] }}</td>
                <td>{{ row['Qtd'] }}</td>
                <td>{{ row['Vl_Unitario'] }}</td>
                <td>{{ row['Preco_Total'] }}</td>
                <td>{{ row['Comissao_Perc'] }}</td>
                <td>{{ row['Valor_Comissao'] }}</td>
                <td class="text-left">{{ row['Cliente'] }}</td>
            </tr>
            {% endfor %}
            <tr class="footer-row">
                <td colspan="6" style="border-right: none; background-color: #fff;"></td>
                <td>{{ total_itens }}</td>
                <td colspan="3" style="text-align: right; padding-right: 15px;">TOTAL:</td>
                <td>R$ {{ total_comissao }}</td>
                <td style="background-color: #fff; border-left: none;"></td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""
template = Template(html_template)

# Configurações do PDF
caminho_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)
opcoes_pdf = {
    'page-size': 'A4',
    'orientation': 'Landscape',
    'margin-top': '15mm',
    'margin-right': '15mm',
    'margin-bottom': '15mm',
    'margin-left': '15mm',
    'enable-local-file-access': None,
    'quiet': ''
}

try:
    string_conexao = (
        f"DRIVER={{{DB_DRIVER}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
    )
    
    conexao = pyodbc.connect(string_conexao)
    df_todas_vendas = pd.read_sql(query, conexao)
    
    if df_todas_vendas.empty:
        print("Nenhuma venda com comissão encontrada no período.")
    else:
        grupos_parceiros = df_todas_vendas.groupby('Cod_Parcei')
        pdfs_gerados = 0
        emails_enviados = 0
        
        for codigo_parceiro, df_vendas in grupos_parceiros:
            total_comissao = df_vendas['Valor_Comissao'].sum()
            
            if total_comissao > 0:
                total_itens = int(df_vendas['Qtd'].sum())
                total_valor_ingressos = df_vendas['Preco_Total'].sum()
                
                nome_parceiro = df_vendas['Parceiro'].iloc[0]
                email_parceiro = df_vendas['Email'].iloc[0]
                
                # Formatando CNPJ ou CPF
                cnpj_raw = str(df_vendas['Cnpj'].iloc[0]).strip()
                if cnpj_raw.endswith('.0'):
                    cnpj_raw = cnpj_raw[:-2]
                    
                if len(cnpj_raw) <= 11:
                    cnpj_raw = cnpj_raw.zfill(11)
                    cnpj_parceiro = f"{cnpj_raw[:3]}.{cnpj_raw[3:6]}.{cnpj_raw[6:9]}-{cnpj_raw[9:]}"
                else:
                    cnpj_raw = cnpj_raw.zfill(14)
                    cnpj_parceiro = f"{cnpj_raw[:2]}.{cnpj_raw[2:5]}.{cnpj_raw[5:8]}/{cnpj_raw[8:12]}-{cnpj_raw[12:]}"
                
                df_display = df_vendas.copy()
                df_display['Qtd'] = df_display['Qtd'].fillna(0).astype(int)
                df_display['Dt_Venda'] = pd.to_datetime(df_display['Dt_Venda']).dt.strftime('%d/%m/%Y')
                
                colunas_moeda = ['Vl_Total', 'Vl_Unitario', 'Preco_Total', 'Valor_Comissao']
                for col in colunas_moeda:
                    df_display[col] = df_display[col].apply(lambda x: f"R$ {x:,.2f}".replace('.', ','))
                df_display['Comissao_Perc'] = df_display['Comissao_Perc'].apply(lambda x: f"{x:,.2f}%".replace('.', ','))
                
                html_renderizado = template.render(
                    logo_base64=logo_base64,
                    nome_parceiro=nome_parceiro,
                    codigo_parceiro=codigo_parceiro,
                    cnpj_parceiro=cnpj_parceiro,   
                    email_parceiro=email_parceiro, 
                    mes_referencia=mes_referencia,
                    data_emissao=data_emissao,
                    itens=df_display,
                    total_itens=total_itens,
                    total_comissao=f"{total_comissao:,.2f}".replace('.', ',')
                )
                
                nome_parceiro_arquivo = nome_parceiro.replace(' ', '_').replace('/', '-').replace('\\', '-')
                nome_arquivo_pdf = f"{PASTA_RELATORIOS}/bordero_{codigo_parceiro}_{nome_parceiro_arquivo}_{nome_mes_anterior.lower()}.pdf"
                
                pdfkit.from_string(html_renderizado, nome_arquivo_pdf, options=opcoes_pdf, configuration=config)
                print(f"✅ PDF gerado: {nome_parceiro} (Cód: {codigo_parceiro})")
                pdfs_gerados += 1
                
                # --- MONTAGEM DO CORPO DO E-MAIL ---
                valor_ingressos_fmt = f"{total_valor_ingressos:,.2f}".replace('.', ',')
                valor_comissao_fmt = f"{total_comissao:,.2f}".replace('.', ',')
                
                assunto_email = f"Boletim de Comissão - {mes_referencia} - {nome_parceiro}"
                corpo_email = f"""
                <p>Olá <strong>{nome_parceiro}</strong>, como vai?</p>
                <p>Segue em anexo o relatório referente à comissão do período de <strong>{data_inicial_fmt}</strong> até <strong>{data_final_fmt}</strong>.</p>
                <p>Abaixo segue um breve resumo:</p>
                <p>
                    1) Quantidade Total de Ingressos: {total_itens}<br>
                    2) Valor Total dos Ingressos: R$ {valor_ingressos_fmt}<br>
                    3) Valor Total da Comissão/Nota Fiscal: R$ {valor_comissao_fmt}
                </p>
                <p>Em caso de dúvidas, veja o arquivo detalhado em anexo.</p>
                <p>Solicitamos que confira os valores e nos envie a Nota Fiscal e a CHAVE PIX até o dia <strong>10/{mes_atual_fmt}</strong> para que o pagamento seja programado.</p>
                <p>A Nota Fiscal deve ser enviada <strong>SOMENTE</strong> para o e-mail <a href="mailto:guias@skyglasscanela.com.br">guias@skyglasscanela.com.br</a>.</p>
                <br>
                <p>Nossos dados devem constar como <strong>TOMADOR DO SERVIÇO</strong>.<br>
                Seguem os dados para emissão da Nota Fiscal:</p>
                <p style="margin-left: 10px;">
                    <strong>Razão social:</strong> Volare Empreendimentos Turísticos Ltda.<br>
                    <strong>CNPJ:</strong> 29.723.868/0002-82<br>
                    <strong>Endereço:</strong> Rua Constante Félix Orsolin, n° 9800 - Caracol - 95684-880 - Canela - RS<br>
                    <strong>INSCRIÇÃO ESTADUAL:</strong> 022/0082529<br>
                    <strong>INSCRIÇÃO MUNICIPAL:</strong> 42434<br>
                    <strong>E-mail:</strong> guias@skyglasscanela.com.br
                </p>
                <br>
                <p><strong>Use o nosso Portal Consignado</strong><br>
                Venda ingressos para o dia a preço normal ou antecipado com desconto, direto pelo celular ou computador!<br>
                O cliente recebe o link de pagamento e os ingressos por e-mail.</p>
                <p><strong>Pagamento:</strong> Semanal, sem necessidade de emitir Nota Fiscal.</p>
                <p><strong>Controle:</strong> Acompanhe vendas e comissões em tempo real.</p>
                <p><strong>Cadastre-se no Portal:</strong> Chame no WhatsApp: (54) 99706-2952.</p>
                <br>
                <p>Atenciosamente<br>
                <strong>Equipe Financeiro Skyglass</strong></p>
                """
                
                email_envio = email_parceiro 
                enviou = enviar_email_com_anexo(email_envio, assunto_email, corpo_email, nome_arquivo_pdf)
                if enviou:
                    print(f"   -> E-mail enviado com sucesso para {email_envio}")
                    emails_enviados += 1
                
        print(f"\nProcesso concluído! {pdfs_gerados} relatórios gerados e {emails_enviados} e-mails enviados.")
        
except Exception as e:
    # --- TRATAMENTO DE ERRO: GERAÇÃO DE LOG E ALERTA PARA A TI ---
    erro_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERRO CRÍTICO NO BORDERO: {str(e)}\n"
    print(erro_msg)
    
    # Salva o log em arquivo .txt dentro da pasta do mês correspondente
    caminho_log = os.path.join(PASTA_RELATORIOS, "erro_execucao.txt")
    with open(caminho_log, "w", encoding="utf-8") as f_log:
        f_log.write(erro_msg)
        
    # Envia o e-mail de alerta para a TI com o log em anexo
    assunto_alerta = f"ALERTA: Falha na Geração de Borderôs de Comissão - {mes_referencia}"
    corpo_alerta = f"""
    <p>Prezada equipe de TI,</p>
    <p>Ocorreu um erro crítico durante a execução automatizada do script de Borderô de Comissões referente a referência <strong>{mes_referencia}</strong>.</p>
    <p><strong>Mensagem de Erro:</strong><br><pre>{str(e)}</pre></p>
    <p>O arquivo de log detalhado encontra-se em anexo nesta mensagem.</p>
    <p>Atenciosamente,<br>Sistema de Automação Skyglass</p>
    """
    
    enviar_email_com_anexo(EMAIL_SUPORTE, assunto_alerta, corpo_alerta, caminho_log)
    print("🚨 E-mail de notificação de erro enviado para a TI.")

finally:
    if 'conexao' in locals():
        conexao.close()