import PySimpleGUI as sg

# Define a função que será executada quando o botão for clicado
def execute_code():
    # Executa o código
    import locale
    import openpyxl
    import win32com.client
    import datetime
    from unidecode import unidecode

    # Obtendo Mês Atual e traduzindo para português
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf-8')
    data_mes_anterior = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    nome_mes_anterior = data_mes_anterior.strftime('%B').capitalize()
    ano_atual = datetime.date.today().year.__str__()

    # Crie uma instância do Excel
    xl_app = win32com.client.GetActiveObject("Excel.Application")

    # Obtém a planilha ativa do Excel
    xl_sheet = xl_app.ActiveSheet

    # Obtém o nome da planilha ativa
    sheet_name = xl_sheet.Name

    # Abre o arquivo Excel usando openpyxl
    workbook = openpyxl.load_workbook('C:\\Users\\felip\\Documents\\Leandro 22\\Fatura de cobranças LEANDRO1.xlsx')
    sheet = workbook[sheet_name]

    # Obtenha o valor da célula K8 na planilha ativa do Excel
    valor_celula = sheet['K8'].value
    namePadrao = valor_celula + " - " + nome_mes_anterior + " " + ano_atual
    namePadrao_SemAcentos = unidecode(namePadrao)

    # Salvar a planilha atual como PDF

    pdf_path = "C:\\Users\\felip\\Documents\\Leandro 23\\{}\\".format(valor_celula) + namePadrao + ".pdf"
    sheet = xl_app.ActiveSheet
    sheet.ExportAsFixedFormat(0, pdf_path, 1, 0)

    # Envio de Email
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders

    # Configuração Email
    if valor_celula == "Aliança Participação":
        para = 'leandrinhofsantos@hotmail.com'
        print(valor_celula)

    elif valor_celula == "Lábios de Mel":
        para = 'faturamento@labiosroupas.com.br'
        print(valor_celula)

    elif valor_celula == "Marlens":
        para = 'aloisio.martins@marlens.com.br'
        print(valor_celula)

    elif valor_celula == "HD":
        para = 'leandrinhofsantos@hotmail.com'
        print(valor_celula)

    elif valor_celula == "Rellus":
        para = 'financeiro@vestme.ind.br'
        print(valor_celula)

    elif valor_celula == "Conceitun":
        para = 'desenvolvimento@conceitun.com.br'
        print(valor_celula)

    elif valor_celula == "Convés":
        para = 'mauricio@convesroupas.com.br'
        print(valor_celula)

    elif valor_celula == "Alpha Lav":
        para = 'leandrinhofsantos@hotmail.com'
        print(valor_celula)

    elif valor_celula == "Mafari":
        para = 'desenho@mafari.com.br'
        print(valor_celula)

    elif valor_celula == "Arielps":
        para = 'arielps@arielps.com.br'
        print(valor_celula)

    elif valor_celula == "Outros":
        para = 'leandrinhofsantos@hotmail.com'
        print(valor_celula)

    assunto = "Relatório de Serviços - " + namePadrao
    mensagem = 'A LFS Transportes agradece a sua preferência!'

    # arquivo a ser enviado
    anexo = pdf_path

    # criar o objeto MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = 'leandrinhofsantos@hotmail.com'
    msg['To'] = para
    msg['Subject'] = assunto

    # adicionar o texto da mensagem
    msg.attach(MIMEText(mensagem, 'plain'))

    # adicionar o anexo ao email
    attachment = open(anexo, "rb")
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= {}.pdf".format(namePadrao_SemAcentos))
    msg.attach(p)

    # enviar o email usando o servidor SMTP do Outlook
    servidor = smtplib.SMTP('smtp.office365.com', 587)
    servidor.starttls()
    servidor.login("leandrinhofsantos@hotmail.com", "****")
    texto = msg.as_string()
    servidor.sendmail("leandrinhofsantos@hotmail.com", para, texto)
    servidor.quit()

    print("finalizado")

    #informa que o código foi concluído
    sg.popup('Código concluído! ' + namePadrao)

# Define a interface gráfica
layout = [
    [sg.Button('Executar código', size=(20, 2))]
]

# Cria a janela
window = sg.Window('Interface com PySimpleGUI').layout(layout)
window.set_icon(icon='C:\\Users\\felip\\PycharmProjects\\pythonProject1\\LFS_AutomateSaveAndSendEmail\\Logo.ico')

# Loop principal de eventos
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Executar código':
        execute_code()

# Fecha a janela ao sair do loop principal
window.close()
