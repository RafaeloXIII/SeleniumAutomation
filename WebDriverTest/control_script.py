import subprocess

scripts = [
    "WebDriverTest/Assinado_SeleniumBot.py",
    "WebDriverTest/Avarias.py",
    "WebDriverTest/consulta_bipagens.py",
    "WebDriverTest/consulta_condutores.py",
    "WebDriverTest/D-1_Seleniumbot.py",
    "WebDriverTest/D2Ddetalhes_seleniumbot.py",
    "WebDriverTest/D2Dresumo_seleniumbot.py",
    "WebDriverTest/falta_pacotesfisicos.py",
    "WebDriverTest/Faltabip_Hub Resumo.py",
    "WebDriverTest/Faltabip_Hub_Detalhes.py",
    "WebDriverTest/FaltaBipeLM_Detalhes.py",
    "WebDriverTest/FaltaBipeLM_Resumo.py",
    "WebDriverTest/GESTÃO DE PACOTES DEVOLVIDOS.py",
    "WebDriverTest/Gestao_pedidosJMS.py",  
    "WebDriverTest/minha_declaracao.py",
    "WebDriverTest/monitoramento_bipagemcoleta.py",
    "WebDriverTest/MonitoramentoBacklog_seleniumbot.py",
    "WebDriverTest/PacotesRetidosLM_seleniumbot.py",
    "WebDriverTest/shippingtime.py",
    "WebDriverTest/Tickets em Aberto.py",
]


for script in scripts:
    try:
        print(f"\033[1;35mSTARTED {script}:\033[0m")
        subprocess.call(["python", script])
    except Exception as e:
        print(f"An error occurred while running {script}: {str(e)}")