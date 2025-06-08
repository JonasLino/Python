import pyautogui
import time
import subprocess
import os
import glob


RA = "2024234029"
downloads_path = "C:\\Users\\francisco.lino\\Downloads"

""" Acessando Valoriza """
subprocess.Popen(["C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"])  
time.sleep(2)

pyautogui.write("https://portalcarreiras.unifametro.edu.br/loginempresa.aspx")
pyautogui.press("enter")
time.sleep(2)

pyautogui.write("suporte.unifametro")
time.sleep(2)
pyautogui.press("enter")


time.sleep(1)
pyautogui.press("enter")

""" ALUNO """
time.sleep(2)
pyautogui.click(x=753, y=194)
time.sleep(1)
pyautogui.moveTo(x=700, y=469)
time.sleep(1)
pyautogui.click(x=437, y=582)
time.sleep(1)
pyautogui.click(x=216, y=462)
time.sleep(1)
pyautogui.click(x=434, y=358)

time.sleep(2)


"""Abre o SQL Server Management Studio e conecta ao servidor."""
def ssms():
    pyautogui.press("win")
    time.sleep(2)
    pyautogui.write("SQL Server Management Studio")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(8)
    pyautogui.press("tab", presses=5)
    pyautogui.press("enter")

def abrir_arquivo_ssms():
    """Abre o arquivo aluno.sql no SSMS."""
    time.sleep(5)
    pyautogui.hotkey("ctrl", "o")
    time.sleep(2)
    pyautogui.write("R:\\NTI-Sistemas\\Lino\\SQL\\Valoriza")  
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.write("aluno.sql")
    pyautogui.press("enter")
    time.sleep(6)

def selecionar_banco():
    """Seleciona o banco Corpore."""
    pyautogui.hotkey("ctrl", "u")
    time.sleep(2)
    pyautogui.write("Corpore")
    pyautogui.press("enter")
    time.sleep(2)

def substituir_ra():
    """Localiza e substitui o RA no script aberto no SSMS."""
    pyautogui.hotkey("ctrl", "h")
    time.sleep(2)
    pyautogui.write("SALUNO.RA = ''")
    pyautogui.press("tab")
    pyautogui.write(f"SALUNO.RA = '{RA}'")
    pyautogui.press("tab", presses=4)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.press("esc")
    pyautogui.press("F5")
    time.sleep(4)
    pyautogui.click(x=125, y=445)
    pyautogui.hotkey("ctrl", "c")

def main():
    """Função principal para executar os scripts de automação."""
    ssms()
    abrir_arquivo_ssms()
    selecionar_banco()
    substituir_ra()

if __name__ == "__main__":
    main()


""" Caminho para a pasta de downloads """
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

files = sorted(
    (os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)),
    key=os.path.getmtime,
    reverse=True
)

if files:
    time.sleep(2)
    os.startfile(files[0])
    time.sleep(2)
else:
    print("Nenhum arquivo encontrado na pasta Downloads.")


""" Posição da célula A2 no Excel """
time.sleep(3)
pyautogui.hotkey('ctrl', 'g')
time.sleep(1)
pyautogui.write('A2')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.hotkey("ctrl", "b")
time.sleep(1)
pyautogui.hotkey("ctrl", "w")
time.sleep(2)
pyautogui.hotkey('alt', 'f4')
time.sleep(2)


""" Retorna ao Site """
pyautogui.keyDown("alt")
pyautogui.press("tab")
pyautogui.keyUp("alt")


""" Importação no Site """
all_files = glob.glob(os.path.join(downloads_path, "*"))
latest_file = max(all_files, key=os.path.getctime) 
time.sleep(3)
pyautogui.click(x=144, y=400)
time.sleep(3)
pyautogui.write(latest_file)
time.sleep(4)
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=146, y=458) #Salvar


""" Caminho para Importação """
time.sleep(1)
pyautogui.click(x=753, y=194)
time.sleep(1)
pyautogui.moveTo(x=700, y=469)
time.sleep(1)
pyautogui.click(x=437, y=582)
time.sleep(2)
pyautogui.write(RA)
time.sleep(2)
pyautogui.click(x=437, y=464)
pyautogui.press("enter") #Integrar


""" CURSO """
time.sleep(2)
pyautogui.click(x=753, y=194)
time.sleep(2)
pyautogui.moveTo(x=700, y=469)
time.sleep(2)
pyautogui.click(x=429, y=612)
time.sleep(2)
pyautogui.click(x=216, y=462)
time.sleep(2)
pyautogui.click(x=567, y=359) #Download

time.sleep(2)

""" Retorna ao SSMS """
pyautogui.keyDown("alt")
pyautogui.press("tab")
pyautogui.keyUp("alt")


def abrir_arquivo_ssms():
    """Abre o arquivo aluno.sql no SSMS."""
    time.sleep(2)
    pyautogui.hotkey("ctrl", "o")
    time.sleep(2)
    pyautogui.write("R:\\NTI-Sistemas\\Lino\\SQL\\Valoriza") 
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.write("curso.sql")
    pyautogui.press("enter")
    time.sleep(2)

def substituir_ra():
    """Localiza e substitui o RA no script aberto no SSMS."""
    pyautogui.hotkey("ctrl", "h")
    time.sleep(2)
    pyautogui.write("SALUNO.RA = ''")
    pyautogui.press("tab")
    pyautogui.write(f"SALUNO.RA = '{RA}'")
    pyautogui.press("tab", presses=4)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.press("esc")
    pyautogui.press("F5")
    time.sleep(4)
    pyautogui.click(x=125, y=445)
    pyautogui.hotkey("ctrl", "c")

def main():
    """Função principal para executar os scripts de automação."""
    abrir_arquivo_ssms()
    substituir_ra()

if __name__ == "__main__":
    main()


""" Caminho para a pasta de downloads """
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

files = sorted(
    (os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)),
    key=os.path.getmtime,
    reverse=True
)

if files:
    time.sleep(2)
    os.startfile(files[0])
    time.sleep(2)
else:
    print("Nenhum arquivo encontrado na pasta Downloads.")


""" Posição da célula A2 no Excel """
time.sleep(3)
pyautogui.hotkey('ctrl', 'g')
time.sleep(1)
pyautogui.write('A2')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.hotkey("ctrl", "b")
time.sleep(1)
pyautogui.hotkey("ctrl", "w")
time.sleep(2)
pyautogui.hotkey('alt', 'f4')
time.sleep(2)

""" Retorna ao Site """
pyautogui.keyDown("alt")
pyautogui.press("tab")
pyautogui.keyUp("alt")


""" Importação no Site """
all_files = glob.glob(os.path.join(downloads_path, "*"))
latest_file = max(all_files, key=os.path.getctime) 
time.sleep(3)
pyautogui.click(x=144, y=400)
time.sleep(3)
pyautogui.write(latest_file)
time.sleep(4)
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=147, y=443) #Salvar


""" Caminho para Importação """
time.sleep(3)
pyautogui.click(x=753, y=194)
time.sleep(1)
pyautogui.moveTo(x=700, y=469)
time.sleep(1)
pyautogui.click(x=429, y=612)
time.sleep(2)
pyautogui.write(RA)
time.sleep(3)
pyautogui.click(x=437, y=464)
pyautogui.press("enter") #Integrar

""" DISCIPLINA """
time.sleep(3)
pyautogui.click(x=753, y=194)
time.sleep(1)
pyautogui.moveTo(x=700, y=469)
time.sleep(1)
pyautogui.click(x=423, y=650)
time.sleep(2)
pyautogui.click(x=216, y=462)
time.sleep(2)
pyautogui.click(x=617, y=358) #Download

time.sleep(4)

""" Retorna ao SSMS """
pyautogui.keyDown("alt")
pyautogui.press("tab")
pyautogui.keyUp("alt")

def abrir_arquivo_ssms():
    """Abre o arquivo aluno.sql no SSMS."""
    time.sleep(2)
    pyautogui.hotkey("ctrl", "o")
    time.sleep(2)
    pyautogui.write("R:\\NTI-Sistemas\\Lino\\SQL\\Valoriza") 
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.write("disciplinas.sql")
    pyautogui.press("enter")
    time.sleep(2)


def substituir_ra():
    """Localiza e substitui o RA no script aberto no SSMS."""
    pyautogui.hotkey("ctrl", "h")
    time.sleep(2)
    pyautogui.write("SALUNO.RA = ''")
    pyautogui.press("tab")
    pyautogui.write(f"SALUNO.RA = '{RA}'")
    pyautogui.press("tab", presses=4)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.press("esc")
    pyautogui.press("F5")
    time.sleep(4)
    pyautogui.click(x=125, y=445)
    pyautogui.hotkey("ctrl", "c")

def main():
    """Função principal para executar os scripts de automação."""
    abrir_arquivo_ssms()
    substituir_ra()

if __name__ == "__main__":
    main()

""" Caminho para a pasta de downloads """
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

files = sorted(
    (os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)),
    key=os.path.getmtime,
    reverse=True
)

if files:
    time.sleep(2)
    os.startfile(files[0])
    time.sleep(2)
else:
    print("Nenhum arquivo encontrado na pasta Downloads.")


""" Posição da célula A2 no Excel """
time.sleep(4)
pyautogui.hotkey('ctrl', 'g')
time.sleep(1)
pyautogui.write('A2')
time.sleep(2)
pyautogui.press('enter')
time.sleep(2)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)
pyautogui.hotkey("ctrl", "b")
time.sleep(2)
pyautogui.hotkey("ctrl", "w")
time.sleep(2)
pyautogui.hotkey('alt', 'f4')
time.sleep(2)

""" Retorna ao Site """
pyautogui.keyDown("alt")
pyautogui.press("tab")
pyautogui.keyUp("alt")

""" Importação no Site """
all_files = glob.glob(os.path.join(downloads_path, "*"))
latest_file = max(all_files, key=os.path.getctime) 
time.sleep(2)
pyautogui.click(x=144, y=400)
time.sleep(2)
pyautogui.write(latest_file)
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=146, y=458) #Salvar

""" Caminho para Importação """
time.sleep(2)
pyautogui.click(x=753, y=194)
time.sleep(1)
pyautogui.moveTo(x=700, y=469)
time.sleep(1)
pyautogui.click(x=429, y=652)
time.sleep(2)
pyautogui.write(RA)
time.sleep(2)
pyautogui.click(x=437, y=464)
pyautogui.press("enter") #Integrar
time.sleep(2)
pyautogui.hotkey('alt', 'f4')

print("Aluno integrado com sucesso")