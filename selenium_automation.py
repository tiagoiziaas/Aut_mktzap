# selenium_automation.py
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from config import mktzap_url, EMAIL, SENHA
from message_reader import ler_paragrafo
from excel_handler import obter_nome_cliente, colorir_linha_verdadeira, obter_nome_vaga  # Adicione `obter_nome_vaga`
import pyautogui
import os

def configurar_driver():
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # Descomente se quiser rodar em modo headless
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(10)
    return driver

def fazer_login(driver, email, senha):
    driver.get(mktzap_url)
    print("Navegador aberto e site carregado")
    try:
        email_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'login_fldLoginUserName'))
        )
        email_input.send_keys(email)

        senha_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'login_fldLoginPassword'))
        )
        senha_input.send_keys(senha)

        login_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'login_btnLogin'))
        )
        login_button.click()
        print("Login realizado, aguardando a próxima página carregar")
    except Exception as e:
        print(f"Erro ao realizar o login: {e}")
        driver.quit()


def navegar_para_secao_financeiro(driver):
    try:
        try:
            elemento = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="sector_36553"]'))
            )
            ActionChains(driver).move_to_element(elemento).perform()
            elemento.click()
            print(f"Elemento clicado:" )
        except Exception as e:
            print('')
        time.sleep(5)
        enviar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@data-test="attendances-button-send_active"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", enviar_button)
        enviar_button.click()
        print("Botão 'Enviar' clicado")

        time.sleep(5)
        whatsapp_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@data-test="attendances-button-send_whatsapp_active"]'))
        )
        whatsapp_button.click()
        print("Opção 'Mensagem no WhatsApp' selecionada")

        select_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'activeChannel'))
        )
        select_element.click()

        canal_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//option[@value="string:waweb_11764"]'))
        )
        canal_button.click()
        try:
            # Encontrar o checkbox
            checkbox = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="responsible"]'))
            )
            time.sleep(5)
            # Rolar até o checkbox
            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
            
            # Usar ActionChains para clicar no checkbox
            actions = ActionChains(driver)
            actions.move_to_element(checkbox).click().perform()
        except Exception as e:
            print(f"Erro ao marcar o checkbox: {e}")
        print("Canal 'Financeiro' selecionado")
    except Exception as e:
        print(f"Erro ao acessar a seção 'Financeiro inicial': {e}")
        driver.quit()

def clicar_card_por_telefone(driver, telefone, index, df, mensagem_file_path, boleto_folder_path, arquivo_excel_path):
    try:
        elemento_span = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{telefone}')]"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(elemento_span).perform()

        elemento_card = elemento_span.find_element(By.XPATH, "./ancestor::div[contains(@class, 'mkz-card-container')]")
        elemento_card.click()

        time.sleep(2)

        nome = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "nome")))
        nome.send_keys(obter_nome_cliente(df, index))

        time.sleep(2)

        msg1 = WebDriverWait(driver, 20).until(EC.visibility_of_element_located(
            (By.XPATH, '//textarea[@id="fieldMessage" and @placeholder="Digite sua mensagem" and contains(@class, "form-control")]')))
        msg1.send_keys(ler_paragrafo(mensagem_file_path, 2))

        time.sleep(3)

        enviar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@type='button' and contains(@class, 'btn-secondary') and contains(@class, 'btn-block') and @ng-click='vm.sendMessage(vm.sector_id)']")))
        enviar.click()
        time.sleep(3)

        msg2 = WebDriverWait(driver, 20).until(EC.visibility_of_element_located(
            (By.XPATH, '//textarea[@id="fieldMessage" and @placeholder="Digite sua mensagem" and contains(@class, "form-control")]')))
        msg2.send_keys(ler_paragrafo(mensagem_file_path, 3))

        time.sleep(2)

        enviar.click()
        time.sleep(3)

        msg3 = WebDriverWait(driver, 20).until(EC.visibility_of_element_located(
            (By.XPATH, '//textarea[@id="fieldMessage" and @placeholder="Digite sua mensagem" and contains(@class, "form-control")]')))
        msg3.send_keys(ler_paragrafo(mensagem_file_path, 4))

        time.sleep(3)
        enviar.click()

        time.sleep(5)

        close = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "close")))
        driver.execute_script("arguments[0].click();", close)

        time.sleep(2)

    except Exception as e:
        print(f"Erro ao clicar no card: {e}")

def enviar_mensagem(driver, index, df, mensagem_file_path, arquivo_excel):
    try:
        telefone = df.iloc[index, 1]
        nome_cliente = obter_nome_cliente(df, index)  # Obtém o nome do cliente
        nome_vaga = obter_nome_vaga(df, index)  # Obtém o nome da vaga

        if not telefone:
            print(f"Telefone não encontrado para o índice: {index}")
            return

        print(f"Enviando mensagem para o telefone: {telefone}")

        # Prepara a mensagem personalizada
        mensagem = ler_paragrafo(mensagem_file_path, nome_cliente, nome_vaga)
        if not mensagem:
            print(f"Erro ao gerar a mensagem para o cliente: {nome_cliente}")
            return

        # Localizar o campo de telefone e inserir o número
        telefone_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@placeholder="(__) _____-____"]'))
        )
        driver.execute_script("arguments[0].click();", telefone_input)
        telefone_input.clear()
        telefone_input.send_keys(telefone)

        # Localizar o campo de mensagem e enviar a mensagem
        mensagem_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'activeMessage'))
        )
        mensagem_input.send_keys(mensagem)

        time.sleep(2)

        print(f"Mensagem enviada para: {nome_cliente} no telefone {telefone}")

        time.sleep(3)
        # Adicionar outros passos se necessário, como clicar no botão "Enviar"
        criar_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Criar") and @type="button" and contains(@class, "btn btn-secondary") and contains(@ng-click, "activeVm.createActiveWhatsapp()")]'))
            )
        driver.execute_script("arguments[0].click();", criar_button)
        print(f"Mensagem enviada para {telefone}")
        colorir_linha_verdadeira(arquivo_excel, index + 2, "FF0000")  # vermelho
        try:
                alerta = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.mkz-toast-warning'))
                )
                if alerta:
                    print("Alerta detectado. Clicando em 'Cancelar'.")
                    cancelar_button = WebDriverWait(driver, 20).until(
                        EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Cancelar") and contains(@class, "btn-outline-primary") and @ng-click="activeVm.close()"]'))
                    )
                    driver.execute_script("arguments[0].click();", cancelar_button)
                    colorir_linha_verdadeira(arquivo_excel, index + 2, "FFFF00")  # amarelo

                    return
                
        except:
            print('Erro')
        

    except Exception as e:
        print(f"Erro ao enviar mensagem: {e}")
    finally:
        time.sleep(5)
