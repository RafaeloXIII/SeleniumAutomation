import subprocess

scripts = [
    "WebDriverRS/Assinado_SeleniumBot.py",
    "WebDriverRS/D-1_Data_Prevista.py",
    "WebDriverRS/D-1_Resumo.py",
    "WebDriverRS/Faltabip_Hub Resumo.py",
    "WebDriverRS/D-1_Seleniumbot.py",
    "WebDriverRS/Faltabip_Hub_Detalhes.py",
    "WebDriverRS/FaltaBipeLM_Detalhes.py",
    "WebDriverRS/FaltabipeLM_Resumo.py",
    "WebDriverRS/Gestao_pedidosJMS.py",
    "WebDriverRS/GestaoRemessa.py",
    "WebDriverRS/minha_declaracao.py",
    "WebDriverRS/MovimentacaoTempoReal.py",
    "WebDriverRS/pacote_problematico.py",
    "WebDriverRS/PacotesRetidosLM_seleniumbot.py",  
    "WebDriverRS/TaxaDevolucao.py",
    "WebDriverRS/TaxaExpedicaoHub.py",
    "WebDriverRS/TaxaExpedicaoHubResumo.py",
    "WebDriverRS/TaxaTransfereciaColeta.py",
]


for script in scripts:
    try:
        print(f"\033[1;35mSTARTED {script}:\033[0m")
        subprocess.call(["python", script])
    except Exception as e:
        print(f"An error occurred while running {script}: {str(e)}")