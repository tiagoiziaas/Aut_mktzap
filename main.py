# main.py
from selenium_automation import configurar_driver, fazer_login, navegar_para_secao_financeiro, enviar_mensagem
from excel_handler import ler_excel
from config import EMAIL, SENHA, mensagem_file_path, arquivo_excel_path
from utils import fehca, preencher_nome, clicar_card_por_numero
import time

def main():
    df = ler_excel()
    if df.empty:
        print("DataFrame vazio. Verifique o arquivo Excel.")
        return
    
    driver = configurar_driver()
    fazer_login(driver, EMAIL, SENHA)

    # Aguarda o login e a página principal carregar
    time.sleep(5)

    for index in range(len(df)):
        try:
            # Você pode descomentar as linhas abaixo se quiser realizar mais ações
            navegar_para_secao_financeiro(driver)
            enviar_mensagem(driver, index, df, mensagem_file_path, arquivo_excel_path)
            # Clica no card correspondente ao número de telefone
            
            time.sleep(5)
            # clicar_card_por_numero(driver)
            time.sleep(5)
            # preencher_nome(driver)
            
        except Exception as e:
            print(f"Erro ao processar a linha {index + 1}: {e}")
    driver.quit()

if __name__ == "__main__":
    main()
