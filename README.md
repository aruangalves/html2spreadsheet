# html2spreadsheet
Parses formatted tables from HTML files into a spreadsheet file

This script reads all HTML files in its folder, parses the relevant data and inserts them into a XLSX file called "PROBEX2019.xlsx" in the same folder
Required data in HTML file:
Código, Título, Categoria (sempre será "Projeto"), Abrangência, Ano (sempre será 2019), Unidade proponente, Unidade orçamentária (sempre será "Não consta"), Outras unidades (sempre será "Não consta"), Área CNPQ, Área principal, Público alvo interno, Público alvo externo, Fonte financiamento, Faz parte de Programa de Extensão (sempre será "NÃO"), Coordenação, Resumo, Metodologia, Objetivos gerais, Fundamentação teórica, Referências
NOTE: This script requires BeautifulSoup and openpyxl modules
AUTHOR: Aruan Galves Nascimento Amaral
CREATED: May 26, 2019
LAST MODIFIED: May 28, 2019
