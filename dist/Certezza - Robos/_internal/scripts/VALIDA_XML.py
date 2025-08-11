import os

# Caminhos
caminho_lista = r"C:\VS_CODE\TRATAR_XML_V2\OURO FINO - 2833 - INSUMOS.txt"
pasta_downloads = r"C:\Users\erison.mendes\Downloads"
arquivo_saida = r"C:\VS_CODE\TRATAR_XML_V2\ourofino_faltantes.txt"

def ler_lista_baixar(caminho):
    with open(caminho, 'r', encoding='utf-8') as f:
        return [linha.strip() for linha in f if linha.strip()]

def listar_arquivos_baixados(pasta):
    arquivos = set()
    for nome in os.listdir(pasta):
        if nome.lower().endswith('.xml'):
            nome_base = os.path.splitext(nome)[0]
            if nome_base.startswith("NFE-"):
                nome_base = nome_base[4:]  # Remove o prefixo "NFE-"
            arquivos.add(nome_base)
    return arquivos

def salvar_faltantes(lista, caminho):
    with open(caminho, 'w', encoding='utf-8') as f:
        for item in lista:
            f.write(item + '\n')
    print(f"\n📝 Arquivo 'faltantes.txt' salvo com {len(lista)} itens.")

def main():
    lista_baixar = ler_lista_baixar(caminho_lista)
    baixados = listar_arquivos_baixados(pasta_downloads)
    
    faltantes = [item for item in lista_baixar if item not in baixados]

    print(f"🔎 Total na lista: {len(lista_baixar)}")
    print(f"✅ Arquivos XML encontrados: {len(baixados)}")
    print(f"⚠️ Arquivos faltantes: {len(faltantes)}")

    if faltantes:
        salvar_faltantes(faltantes, arquivo_saida)
    else:
        print("🎉 Todos os arquivos foram encontrados.")

if __name__ == "__main__":
    main()
