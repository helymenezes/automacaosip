import pandas as pd
import pyautogui
import time

        
path_excel = r'mla_emt_2t.xlsx'

# Preenchimento de UC(unidade consumidora / obra) do sistema Gestão SIP One Engenharia
cabecalho = 9
plan_instalacao = pd.read_excel(path_excel, header= cabecalho)
df = plan_instalacao[['Código ONE', 'Uc', 'Nº Obra']]
df = df.replace(0, pd.NA).dropna().dropna(axis=1)
df['Nº Obra'] = df['Nº Obra'].astype(str).str.rstrip('.0')

#Alerta para inicio:
pyautogui.alert("Aplicação esta funcionando neste nomento não interropa o computador")
pyautogui.PAUSE = 0.5

# Acessando loguin
pyautogui.press('winleft')
pyautogui.write('gestao')#SIP
time.sleep(2)
pyautogui.press('enter')
time.sleep(4)
pyautogui.write('hely')
pyautogui.press('tab') 
pyautogui.write('hrs500')
time.sleep(2.5)
pyautogui.moveTo(999,488)
pyautogui.click(clicks = 1, button = 'left')
time.sleep(3.5)

#Acessando o Painel Relatório:
pyautogui.moveTo(505,49)
pyautogui.click(clicks = 1, button = 'left')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')

# Realizando  a busca do cod UC e Numero de obrhely hrs500a:
time.sleep(1)
#print(pyautogui.position())
pyautogui.moveTo(x=959, y=278, duration=1.5)
pyautogui.click(clicks = 1, button = 'left')

# Armazenar a posição inicial do mouse para retorno
start_position = pyautogui.position()

for i, row in df.iterrows():
    codigo_one = str(row['Código ONE'])
    codigo_uc = str(row['Uc'])
    codigo_obra = row['Nº Obra']
    
    if i == len(df) - 1:
        pyautogui.alert(f"O último número do Código ONE é {codigo_one}")
        break
    
    time.sleep(2)
    pyautogui.write(codigo_one)
    time.sleep(2)
    pyautogui.moveTo(x=1233, y=265, duration=1.5)
    pyautogui.click(clicks = 1, button = 'left')
    time.sleep(2)
    pyautogui.moveTo(x=99, y=562, duration= 1.5)
    time.sleep(2)
    pyautogui.click(clicks = 1, button = 'left')
    time.sleep(2)
    pyautogui.moveTo(x=655, y=87, duration= 1.5)
    time.sleep(2)
    pyautogui.click(clicks = 1, button = 'left')
    time.sleep(2)
    pyautogui.moveTo(x=1132, y=144, duration= 1.5)
    pyautogui.click(clicks = 1, button = 'left')
    pyautogui.write(codigo_uc)
    time.sleep(2)
    pyautogui.moveTo(x=762, y=117, duration= 1.5)
    time.sleep(2)
    pyautogui.click(clicks = 1, button = 'left')
    pyautogui.write(codigo_obra)
    time.sleep(2)
    pyautogui.moveTo(x=738, y=89, duration= 1.5)
    pyautogui.click(clicks = 1, button = 'left')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.moveTo(x=951, y=79, duration= 1.5)
    pyautogui.click(clicks = 1, button = 'left')
    #print(pyautogui.position())
    #Retornar à posição inicial
    pyautogui.moveTo(start_position)
    pyautogui.click(clicks = 2, button = 'left')
    pyautogui.press('del')
    

    

    


    # Limpando o campo de entrada para a próxima iteração
    # pyautogui.moveTo(x=959, y=278, duration=1)
    # pyautogui.click(clicks=1, button='left')
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.press('backspace')
    # time.sleep(2)