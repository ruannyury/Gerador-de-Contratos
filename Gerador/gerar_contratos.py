import pandas as pd
from docx import Document
import os

# Função para carregar dados e preencher contratos
def gerar_contratos(planilha_caminho, modelo_caminho, pasta_destino="Contratos", pasta_incompletos="Contratos_Incompletos"):
    # Carregar a planilha de dados
    dados = pd.read_excel(planilha_caminho)

    # Criar pastas de destino, se não existirem
    os.makedirs(pasta_destino, exist_ok=True)
    os.makedirs(pasta_incompletos, exist_ok=True)

    # Função para preencher o contrato
    def preencher_contrato(dados, modelo_caminho):
        # Carrega uma nova instância do modelo para cada contrato
        contrato = Document(modelo_caminho)
        for paragrafo in contrato.paragraphs:
            for campo in ['Name', 'Valor Hora Aula', 'E-mail', 'CPF', 'RG', 'Orgão Expedidor',
                          'Data de Nascimento', 'Endereço', 'Número', 'Bairro', 'Complemento',
                          'Cidade', 'Estado (UF)', 'CEP', 'Telefone 1', 'Nacionalidade',
                          'Estado Civil', 'CNPJ', 'Nome empresarial', 'Endereço da sede',
                          'Representante legal', 'Dados Bancários']:
                valor = dados.get(campo, "")
                # Verifica se o valor está vazio e substitui por "(Campo FALTANTE)"
                if pd.isna(valor) or valor == "":
                    substituto = f"({campo} FALTANTE)"
                else:
                    substituto = str(valor)
                paragrafo.text = paragrafo.text.replace(f'[{campo}]', substituto)
        return contrato

    # Loop para gerar um contrato para cada professor
    for _, linha in dados.iterrows():
        dados_professor = linha.to_dict()

        # Verifica campos vazios e cria uma lista com os campos faltantes
        campos_faltantes = [campo for campo in ['Name', 'Valor Hora Aula', 'E-mail', 'CPF', 'RG', 'Orgão Expedidor',
                                                'Data de Nascimento', 'Endereço', 'Número', 'Bairro', 'Complemento',
                                                'Cidade', 'Estado (UF)', 'CEP', 'Telefone 1', 'Nacionalidade',
                                                'Estado Civil', 'CNPJ', 'Nome empresarial', 'Endereço da sede',
                                                'Representante legal', 'Dados Bancários']
                            if pd.isna(dados_professor[campo]) or dados_professor[campo] == ""]

        if campos_faltantes:
            # Se há campos faltantes, define o nome do arquivo com a indicação de dados ausentes
            campos_faltantes_str = "_".join(campos_faltantes)
            nome_arquivo = f'{pasta_incompletos}/Contrato_{dados_professor.get("Name", "Desconhecido")}_FALTANDO_{campos_faltantes_str}.docx'
        else:
            # Se não há campos faltantes, salva normalmente
            nome_arquivo = f'{pasta_destino}/Contrato_{dados_professor["Name"]}.docx'

        # Preenche e salva o contrato
        contrato_preenchido = preencher_contrato(dados_professor, modelo_caminho)
        contrato_preenchido.save(nome_arquivo)

        if campos_faltantes:
            print(f'Contrato com dados faltantes gerado: {nome_arquivo}')
        else:
            print(f'Contrato gerado: {nome_arquivo}')


# Caminho dos arquivos (planilha e modelo) e execução do script
if __name__ == "__main__":
    planilha_caminho = "professores.xlsx"  # Caminho para a planilha de dados
    modelo_caminho = "modelo_contrato.docx"  # Caminho para o modelo do contrato
    gerar_contratos(planilha_caminho, modelo_caminho)
