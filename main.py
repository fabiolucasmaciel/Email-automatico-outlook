# Para conectar ao outlook e ler CSVs
import win32com.client as win32
import csv

# integração com o outlook
outlook = win32.Dispatch('outlook.application')

# Abre csv, delimitado por ';'
reader = csv.reader(open('Emails.csv', 'r'), delimiter=';')
for linha in reader:
    # Pega info do csv por partes
    cpf_reg = linha[0]
    nome_reg = linha[1]
    email_reg = linha[2]
    cadastro_reg = linha[3]
    faixa_reg = linha[4]
    nasc_reg = linha[5]
    sexo_reg = linha[6]

    # Definindo dstancia,preço,data e horas pelas info de registro
    if faixa_reg == "Jovem" or faixa_reg == "Sênior":
        distancia = "21km, Meia-Maratona"
        preco = 'R$90'
        if faixa_reg == "Jovem":
            data = "26/11/2021"
            hora_max = "15:00"
            hora_largada = "15:30"
        else:
            data = "26/11/2021"
            hora_max = "07:00"
            hora_largada = "07:30"
    else:
        distancia = "42km, Maratona"
        preco = 'R$120'
        data = "27/11/2021"
        hora_max = "07:00"
        hora_largada = "07:30"

    # Para o usuario do programa ver oque está acontecendo
    print("--------------------------------------------")
    print("Enviando para " + linha[0])
    print("E-mail registrado: " + linha[1])

    # Cria um email novo para enviar
    email = outlook.CreateItem(0)

    # Email destindo definido
    email.To = email_reg
    # Assunto do Email
    email.Subject = "Informações da Maratona Carioca"

    # Texto do corpo do Email
    email.HTMLBody = f"""
        <h1>Maratona CARIOCA 2021</h1>
        <hr>
        
        <h2>Olá, {nome_reg}</h2>
        <h3>Aqui temos algumas informações sobre seu cadastro na maratona</h3>
        <p>Nome de registro: {nome_reg}</p>
        <p>CPF de registro: {cpf_reg}</p>
        <p>Categoria: {faixa_reg} {sexo_reg}</p>
        <p>Número do corredor: {cadastro_reg}</p>
        <p>Distância: {distancia}</p>
        <p>Data da corrida: {data}</p>
        <p>Hora máxima de registro: {hora_max}</p>
        <p>Hora da largada: {hora_largada}</p>
        <p>Data de Nascimento: {nasc_reg}</p>
        <p>Preço Pago: {preco}</p>
        <p>E-mail registrado: {email_reg}</p>
        <a href="google.com">SAIBA MAIS NO SITE</a>
        <hr>
        
        <p>abs,</p>
        <p>CORRIDA CARIOCA 2021</p>
    """

    # Envio do Email
    email.Send()
    print("Email enviado com sucesso")
    print("--------------------------------------------")
print("\nEnvios Concluidos")
