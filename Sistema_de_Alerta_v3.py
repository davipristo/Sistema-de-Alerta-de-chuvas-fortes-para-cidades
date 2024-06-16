import requests
import win32com.client as win32

#API OPEN WHEATER
API_KEY = '0ef015b8881e07d70e1bf4322b77d20b'
CITY = 'Caxias do Sul'
THRESHOLD = 2  # Limite de chuva forte em mm

# PUXAR DADOS PELA API
def get_weather_forecast(api_key, city):
    url = f"http://api.openweathermap.org/data/2.5/forecast?q={city}&appid={api_key}&units=metric"
    print(f"Requisição para URL: {url}")
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro na requisição: {response.status_code}")
        print("Resposta da API:", response.json())
        return None

#VER SE TEM CHUVA
def check_heavy_rain(forecast, threshold):
    if forecast is None:
        return []
    if 'list' not in forecast:
        print("Resposta da API não contém a chave 'list'.")
        print("Resposta completa:", forecast)
        return []

    alerts = []
    for entry in forecast['list']:
        rain = entry.get('rain', {}).get('3h', 0)
        if rain >= threshold:
            alert_time = entry['dt_txt']
            alerts.append(f"Chuva forte de {rain} mm prevista para {alert_time}")
    return alerts

#SEND MAIL
def send_email_with_weather_alerts(alerts):
    # Configurações do email
    sender_email = "davi.pristo@outlook.com"  # Substitua pelo seu email
    recipients = [
        "davi.pristo@outlook.com",
        "matheus.pereira@aedb.br",
        "joao.carneiro@aedb.br",
        "thiagode.pereira@aedb.br",
        "caio.avelar@aedb.br"
    ]
    subject = "Alerta Automático de Chuva Forte"

    #REDIGINDO MAIL
    body = f"""
    <p>Previsão do tempo para {CITY}:</p>
    """
    if alerts:
        body += "<p>Alertas de Chuva Forte:</p>"
        for alert in alerts:
            body += f"<p>{alert}</p>"
        body += """
        <p><strong>Nota:</strong> Organizações de defesa civil devem preparar a cidade e os sistemas de drenagem para as chuvas fortes.</p>
        """
    else:
        body += "<p>Nenhuma chuva forte prevista nos próximos dias.</p>"

    #AVISO MENSAGEM AUTOMATICA
    body += "<p>Este é um alerta automático.</p>"

    #OUTLOOK WIN 32
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = ";".join(recipients)
    email.Subject = subject
    email.HTMLBody = body

    #IF DE ERRO DE ENVIO
    try:
        email.Send()
        print("Email enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar o email: {e}")

def main():
    #Obtendo a previsão do tempo
    forecast = get_weather_forecast(API_KEY, CITY)

    #Alertas de chuva forte
    alerts = check_heavy_rain(forecast, THRESHOLD)

    #print no terminal
    print(f"Previsão do tempo para {CITY}:")
    if alerts:
        print("Alertas de Chuva Forte:")
        for alert in alerts:
            print(alert)
    else:
        print("Nenhuma chuva forte prevista nos próximos dias.")

    #Enviando email com os alertas
    send_email_with_weather_alerts(alerts)

if __name__ == "__main__":
    main()
