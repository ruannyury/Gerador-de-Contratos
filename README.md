# Gerador de Contratos

Este projeto é um gerador de contratos para professores, utilizando dados de uma planilha Excel. O programa preenche um modelo de contrato com informações dos professores e gera documentos Word.

## Pré-requisitos

Antes de executar o programa, certifique-se de ter o Python instalado em seu sistema. Você pode verificar se o Python está instalado executando o seguinte comando no terminal ou prompt de comando:

```bash
python --version
```

ou

```bash
python3 --version
```

Se o Python não estiver instalado, você pode baixá-lo e instalá-lo a partir do [site oficial do Python](https://www.python.org/downloads/).

### Arquivos Necessários

Para que o programa funcione corretamente, você precisa dos seguintes arquivos na pasta do projeto:

- **Planilha de Dados:** O nome do arquivo da planilha deve ser **`professores.xlsx`**.
- **Modelo de Contrato:** O nome do arquivo do modelo do contrato deve ser **`modelo_contrato.docx`**.
- **Arquivo de Dependências:** O arquivo **`requirements.txt`** já deve estar incluído na pasta.

## Instalação das Dependências

O projeto já contém um arquivo `requirements.txt` que lista todas as bibliotecas necessárias para o funcionamento do programa. Para instalar as dependências, siga os passos abaixo:

### Instalando as Dependências

#### No Linux

1. Abra o terminal e navegue até o diretório do seu projeto.

2. Execute o seguinte comando:

   ```bash
   pip3 install -r requirements.txt
   ```

#### No Windows

1. Abra o Prompt de Comando (CMD) e navegue até o diretório do seu projeto.

2. Execute o seguinte comando:

   ```bash
   pip install -r requirements.txt
   ```

## Criar e Ativar um Ambiente Virtual (Opcional)

Recomenda-se o uso de um ambiente virtual para evitar conflitos entre diferentes projetos. Siga os passos abaixo para criar e ativar um ambiente virtual.

### Criar um Ambiente Virtual

1. No terminal ou CMD, navegue até o diretório do seu projeto.

2. Crie um ambiente virtual com o seguinte comando:

   ```bash
   python -m venv venv
   ```

### Ativar o Ambiente Virtual

- **No Linux:**

  ```bash
  source venv/bin/activate
  ```

- **No Windows:**

  ```bash
  venv\Scripts\activate
  ```

Após ativar o ambiente virtual, instale as dependências novamente usando o comando de instalação do `requirements.txt`.

## Executando o Programa

1. Certifique-se de que a planilha de dados (`professores.xlsx`) e o modelo de contrato (`modelo_contrato.docx`) estão na mesma pasta que o script `gerar_contratos.py`.

2. No terminal ou CMD, execute o script com o seguinte comando:

   ```bash
   python gerar_contratos.py
   ```

ou

```bash
python3 gerar_contratos.py
```

O programa irá gerar contratos preenchidos com as informações da planilha e salvá-los nas pastas **`Contratos`** e **`Contratos_Incompletos`**, dependendo da presença de dados faltantes.

## Conclusão

Agora você está pronto para usar o gerador de contratos! Se tiver alguma dúvida ou encontrar problemas, sinta-se à vontade para consultar a documentação do Python ou abrir uma nova questão.
