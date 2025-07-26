import xlwings as xw
import pandas as pd
import glob
import os
import time

# CONFIGS
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
nome_base_arquivo = "Controle de Patio x data ingreso do pedido"
nome_aba_destino = "BASE"

# Caminho destino final
pasta_final = os.path.join(
    os.path.expanduser("~"),
    r"Arcor\GC_CPS_CDs_Controle_de_Patio - Documentos\CONTROLE DE PÁTIO - 2025"
)
padrao_arquivo_final = os.path.join(pasta_final, "?? - CONTROLE DE PÁTIO - * - 2025.xlsm")


def achar_arquivo_mais_novo(padrao):
    lista_arquivos = glob.glob(padrao)
    if not lista_arquivos:
        return None
    return max(lista_arquivos, key=os.path.getmtime)


def esperar_planilha_fechar(caminho_arquivo):
    print("⏳ Verificando se a planilha está aberta...")
    aviso_exibido = False
    while True:
        try:
            with open(caminho_arquivo, "r+", encoding="utf-8") as f:
                break
        except PermissionError:
            if not aviso_exibido:
                print("⚠️ A planilha está aberta no Excel. Feche ela pra continuar...")
                aviso_exibido = True
            time.sleep(2)


def ler_planilha_com_xlwings(caminho_arquivo):
    app = xw.App(visible=False)
    wb = app.books.open(caminho_arquivo)

    try:
        sht = wb.sheets[0]
        valores = sht.used_range.value

        if valores is None or not isinstance(valores, list):
            print("⚠️ A planilha está vazia ou mal formatada.")
            return pd.DataFrame()

        if isinstance(valores[0], str):  # Só uma linha
            valores = [valores]

        header_index = None
        for i, linha in enumerate(valores):
            if isinstance(linha, list) and "Carregamento" in linha:
                header_index = i
                break

        if header_index is None:
            print("❌ Cabeçalho com 'Carregamento' não encontrado.")
            return pd.DataFrame()

        colunas = valores[header_index]
        dados = valores[header_index + 1:]

        dados_validos = [linha for linha in dados if isinstance(linha, list) and len(linha) == len(colunas)]
        if not dados_validos:
            print("❌ Nenhuma linha de dados válida.")
            return pd.DataFrame()

        df = pd.DataFrame(dados_validos, columns=colunas)

        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

        return df

    finally:
        wb.close()
        app.quit()


def atualizar_planilha():
    print("🔍 Procurando arquivos...")
    arquivo_base = achar_arquivo_mais_novo(os.path.join(downloads_path, nome_base_arquivo + "*.xlsx"))
    arquivo_final = achar_arquivo_mais_novo(padrao_arquivo_final)

    if not arquivo_base:
        print("❌ Nenhum arquivo base encontrado na pasta Downloads!")
        input("Pressione Enter para sair.")
        return

    if not arquivo_final:
        print("❌ Nenhum arquivo final encontrado na pasta de destino!")
        input("Pressione Enter para sair.")
        return

    print(f"📥 Arquivo base: {os.path.basename(arquivo_base)}")
    print(f"📂 Arquivo final: {os.path.basename(arquivo_final)}")

    esperar_planilha_fechar(arquivo_final)

    print("📖 Lendo arquivo base...")
    nova_base = ler_planilha_com_xlwings(arquivo_base)
    if nova_base.empty:
        print("❌ Dados inválidos ou planilha vazia.")
        input("Pressione Enter para sair.")
        return

    print("✏️ Atualizando planilha final...")
    app = xw.App(visible=False)
    wb = app.books.open(arquivo_final)

    try:
        sht = wb.sheets[nome_aba_destino]
        sht.clear_contents()
        sht.range("A1").options(index=False).value = nova_base
        wb.save()
        print("✅ Planilha final atualizada com sucesso!")
    finally:
        wb.close()
        app.quit()

    xw.App(visible=True, add_book=False).books.open(arquivo_final)
    input("🚀 Processo concluído! Pressione Enter para fechar.")


if __name__ == "__main__":
    atualizar_planilha()