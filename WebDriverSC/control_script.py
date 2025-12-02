import subprocess

scripts = [
    "WebDriverSC/Assinado_SC_SeleniumBot.py",
    "WebDriverSC/D-1_SC_Seleniumbot.py",
]


for script in scripts:
    try:
        print(f"\033[1;35mSTARTED {script}:\033[0m")
        subprocess.call(["python", script])
    except Exception as e:
        print(f"An error occurred while running {script}: {str(e)}")