
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl
# Entrar no site https://pje-consulta-publica.tjmg.jus.br/
numero_oab = input('Digite o número OAB: ')
estado = input('Digite a sigla do Estado em letras maiúsculas(RJ,SP): ')
nome_advogado = input('Digite o nome do advogado: ')
planilha_dados = openpyxl.load_workbook('dados da consulta.xlsx')
pag_processos = planilha_dados['Planilha1']
sleep(30)
driver = webdriver.Chrome()

driver.get('https://pje-consulta-publica.tjmg.jus.br/')


# Clicar no campo oab e digitar o número do advogado
sleep(3)
campo_oab = driver.find_element(
    By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']")
sleep(1)
campo_oab.click()
sleep(1)
campo_oab.send_keys(numero_oab)

# Selecionar o estado daquele advogado
sleep(1)
campo_uf = driver.find_element(
    By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
sleep(1)
opcoes_uf = Select(campo_uf)
sleep(1)
opcoes_uf.select_by_visible_text(estado)
sleep(1)
# Clicar em pesquisar
botao_pesquisar = driver.find_element(
    By.XPATH, "//input[@id='fPP:searchProcessos']")
sleep(1)
botao_pesquisar.click()
sleep(3)
# Entrar em cada um dos processos e extrair o número do advogado, o número do processo e o nome dos participantes

botoes_detalhes = driver.find_elements(By.XPATH, "//a[@title='Ver Detalhes']")

for botao in botoes_detalhes:
    janela_principal = driver.current_window_handle
    botao.click()
    sleep(5)
    janelas_abertas = driver.window_handles
    for janela in janelas_abertas:
        if janela not in janela_principal:
            driver.switch_to.window(janela)
            sleep(5)
            numero_processo = driver.find_elements(By.XPATH, "//div[@class='propertyView ']//div[@class='col-sm-12 ']")[0]
            nome_participante = driver.find_elements(By.XPATH, "//tbody[contains(@id, 'processoPartesPoloAtivoResumidoList:tb')]//span[@class= 'text-bold']")
            lista_participantes = []
            for participante in nome_participante:
                lista_participantes.append(participante.text)
            if len(lista_participantes)==1:
                pag_processos.append([numero_oab,numero_processo.text,lista_participantes[0]])
            else:
                pag_processos.append([numero_oab,numero_processo.text,','.join(lista_participantes)])
            # Salvar os dados na planilha
            planilha_dados.save(f'dados da consulta {nome_advogado}.xlsx')
            driver.close()
            # Repetir até finalizar todo so processos do advogado
    driver.switch_to.window(janela_principal)

