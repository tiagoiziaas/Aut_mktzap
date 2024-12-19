def ler_paragrafo(mensagem_file_path, nome_cliente, nome_vaga):
    try:
        # Lê o conteúdo do arquivo
        with open(mensagem_file_path, 'r', encoding='utf-8') as file:
            conteudo = file.read()

        # Verifica se o arquivo está vazio
        if not conteudo:
            raise ValueError("O arquivo de mensagem está vazio ou não pôde ser lido.")

        # Substitui os marcadores pelo nome do cliente e da vaga
        mensagem_personalizada = conteudo.replace("{NOME}", nome_cliente).replace("{NOME DA VAGA}", nome_vaga)

        return mensagem_personalizada

    except Exception as e:
        print(f"Erro ao ler ou processar o arquivo de mensagem: {e}")
        return ""
