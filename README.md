# Case_Suzano

# Classificação de Indicadores Ambientais

Este projeto em Python automatiza a classificação de indicadores ambientais em uma planilha Excel. Ele usa a **API de inteligência generativa do Google** para categorizar cada indicador e adiciona uma nova coluna de classificação na planilha original, mantendo a formatação das colunas existentes.

## Funcionalidades

- **Classificação Automática dos Indicadores:**  
  Cada indicador na planilha é classificado pela API como "Água", "Energia" ou "Outros".

- **Manutenção da Formatação da Planilha:**  
  A nova coluna "Classificação" é adicionada com o cabeçalho formatado com **fundo preto**, **fonte branca**, **centralização**, **borda cinza** e **largura da coluna ajustada** para 35,11, de modo a manter a consistência visual com as demais colunas.

## Tecnologias Utilizadas

- **Python**
- **openpyxl** para manipulação de planilhas Excel.
- **google-generativeai** para acesso à API de classificação.
- **pandas** para leitura inicial da planilha.

## Estrutura do Código

1. **Leitura e Carregamento da Planilha Original:**  
   O código carrega a planilha Excel com os indicadores ambientais.

2. **Classificação dos Indicadores:**  
   Para cada indicador, é feita uma requisição à API para classificá-lo. Se a resposta for inválida ou houver erro, a classificação padrão é "Erro".

3. **Adição e Formatação da Nova Coluna:**  
   A coluna "Classificação" é adicionada com o cabeçalho e as configurações de estilo.

4. **Salvamento da Planilha Modificada:**  
   O arquivo Excel é salvo no caminho especificado.
