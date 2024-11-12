# Case_Suzano

Este script em Python tem como objetivo classificar indicadores ambientais presentes em uma planilha Excel, adicionando uma nova coluna de classificação ("Água", "Energia" ou "Outros") baseada em uma API do Google chamada Gemini. Além disso, ele mantém a formatação da planilha original, criando a nova coluna "Classificação" com o mesmo estilo e características das outras colunas existentes. O código utiliza a biblioteca openpyxl para manipulação da planilha e a google.generativeai para acessar a API de IA.

Funcionalidades
Carregar a Planilha Original: O script carrega a planilha Excel existente que contém os dados dos indicadores.

Classificação dos Indicadores: Para cada indicador na planilha, o script envia uma solicitação à API de IA para classificar o indicador como "Água", "Energia" ou "Outros". Caso a classificação não seja válida ou haja um erro, o script registra a mensagem de erro.

Adição de Coluna de Classificação: O script cria uma nova coluna "Classificação" na planilha original, mantendo o cabeçalho e a formatação da planilha.

Formatação de Células: A nova coluna "Classificação" recebe formatação consistente com as outras colunas da planilha:
O cabeçalho é preenchido com a cor preta e a fonte é configurada para branca.
A coluna é centralizada.
A largura da coluna é ajustada para 35,11.
As bordas da célula do cabeçalho são configuradas com a cor #CCCCCC (cinza claro).

Bibliotecas Utilizadas

openpyxl: Biblioteca utilizada para manipulação de planilhas Excel no formato .xlsx.

google.generativeai: API de inteligência generativa do Google utilizada para classificar os indicadores.

Como Funciona

Leitura da Planilha: O script começa lendo a planilha Excel original com a ajuda da biblioteca pandas, carregando as informações necessárias.

Interação com a API do Google: Para cada indicador presente na coluna "Nome do Indicador", uma solicitação é enviada à API do Google para classificá-lo. A classificação pode ser "Água", "Energia" ou "Outros".

Criação da Nova Coluna: Uma nova coluna "Classificação" é criada ao lado da coluna "Nome do Indicador". O cabeçalho da nova coluna é formatado com a cor preta e fonte branca, as bordas das células são aplicadas e a largura da coluna é ajustada.

Salvamento da Planilha Modificada: O script salva a planilha original com a nova coluna "Classificação" e formatação, garantindo que todas as alterações sejam preservadas.
