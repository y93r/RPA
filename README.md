# RPA
##### RPA_ERP
Criação de uma RPA cadastrando automaticamente informações de produtos em uma ERP:
- Biblioteca pyautogui;
- Biblioteca subprocess para abrir o ERP, no caso o Fakturama;
- Uso de funções para encontrar imagem, escrever texto e posição da imagem;
- Preenchimento automático de produtos.

##### rpa_email
Automatizar a extração de informações de um sistema e envio de um relatório por e-mail
- Utilização da biblioteca pyautogui;
- Criar um alerta para não mexer no computador enquanto o código está rodando;
- Tirar print das imagens dos lugares onde você vai clicar;
- Entrar no sistema;
- Entrar em uma aba específica do sistema onde tem o relatório;
- Exportar o Relatório;
- Pegar o relatório exportado, tratar e pegar as informações que queremos com pandas;
- Preencher/Atualizar informações do sistema.

##### Bs4
Web scraping da tabela de criptomoedas do site CoinMarketCap:
- Biblioteca requests para fazer solicitação HTTP;
- Biblioteca BeautifulSoup para fazer a análise do código HTML e encontrar as informações desejadas;
- Extração dos dados que serão armazenados em um dicionário;
- Criar uma tabela com pandas com as informações do dicionário, transformando o índice em coluna;
- Salvar essa tabela no formato excel com a data de hoje
