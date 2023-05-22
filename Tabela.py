import pandas as pd
from selenium import webdriver
import pyautogui as pg
import time

navegador = webdriver.Chrome()
#dados_excel = pd.read_csv("Clientes.csv")
magalu = "https://www.magazineluiza.com.br/".replace('"',"")
magalu_x = '//*[@id="input-search"]'.replace("'","")
amazon = "https://www.amazon.com.br/ ".replace('"',"")
amazon_x = '//*[@id="twotabsearchtextbox"]'

links = [magalu,amazon]
searchbar = [magalu_x,amazon_x]

for i in range(len(links)):
    navegador.get(f"{links[i]}")
    botao = navegador.find_element("xpath", f"{searchbar[i]}")
    botao.click()
    botao.send_keys("S22 Ultra")
    pg.press("enter")
    time.sleep(3)
#pg.press("win")

