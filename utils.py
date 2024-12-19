from selenium_automation import WebDriverWait, By,EC
def fehca(driver):
    try:
        # Aguarda o botão de fechar estar visível e clicável
        botao_fechar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="dlgChatModal"]//button[contains(@class, "close")]'))
        )
        
        # Clicar no botão de fechar
        driver.execute_script("arguments[0].click();",  botao_fechar)
        print("Modal fechado com sucesso.")
    except Exception as e:
        print(f"Erro ao fechar o modal: {e}")

def clicar_card_por_numero(driver):
    try:
        # Localiza o primeiro card com a classe correspondente
        card = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "attendance-card-container")]'))
        )

        # Rola até o card, caso necessário
        driver.execute_script("arguments[0].scrollIntoView(true);", card)

        # Clica no card
        card.click()
        print("Primeiro card clicado com sucesso.")
    except Exception as e:
        print(f"Erro ao clicar no card: {e}")


def preencher_nome(driver):
    """
    Preenche o campo 'nome' no formulário com o valor fornecido.
    
    Args:
        driver: Instância do WebDriver.
        nome_cliente: Nome do cliente extraído da coluna A do Excel.
    """
    try:

        dropdown = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "status_id"))
)

        # Seleciona a opção "PRÉ-SELECIONADO" pelo valor
        dropdown.find_element(By.XPATH, '//option[@value="string:73362"]').click()

        print("Opção 'PRÉ-SELECIONADO' selecionada com sucesso.")
        down = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'modal_confirmation_status_confirm')))
        down.click()
    except Exception as e:
        print(f"Erro ao preencher o campo 'nome': {e}")