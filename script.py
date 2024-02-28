#This script reads all HTML files in its folder, parses the relevant data and
#inserts them into a XLSX file called "PROBEX2019.xlsx" in the same folder
#Required data in HTML file:
#Código, Título, Categoria (sempre será "Projeto"), Abrangência, Ano (sempre será 2019), Unidade proponente, Unidade orçamentária (sempre será "Não consta"), Outras unidades (sempre será "Não consta"), Área CNPQ, Área principal, Público alvo interno, Público alvo externo, Fonte financiamento, Faz parte de Programa de Extensão (sempre será "NÃO"), Coordenação, Resumo, Metodologia, Objetivos gerais, Fundamentação teórica, Referências
#NOTE: This script requires BeautifulSoup and openpyxl modules
#AUTHOR: Aruan Galves Nascimento Amaral
#CREATED: May 26, 2019
#LAST MODIFIED: May 28, 2019

import os
import re
import openpyxl
from bs4 import BeautifulSoup

def formatString(str):
  if(str):
    str = str.replace('\n','').replace('\t','').replace('–','-').replace("","").replace("","")
    #Checks if field is an empty string or has an empty field '-' placeholder
    pattern = re.match(r'^\s*\-*\s*$', str)    
    if(pattern):
      return "Não consta"
    else:
      return str
  else:
    return "Não consta"

path = os.path.dirname(os.path.realpath(__file__))
files = os.listdir(path)
print("Opening excel file: " +path +'\\' +'PROBEX2019.xlsx')
xlsxFile = openpyxl.load_workbook(path + '\\' +'PROBEX2019.xlsx')
sheet = xlsxFile.get_active_sheet()
cellIndex = 2

xlsxHeader = ["Código", "Título", "Categoria", "Abrangência", "Ano", "Unidade proponente", "Unidade orçamentária", "Outras unidades", "Área CNPq", "Área principal", "Público alvo interno", "Público alvo externo", "Fonte financiamento", "Faz parte de Programa de Extensão", "Coordenação", "Resumo", "Metodologia", "Objetivos gerais", "Fundamentação teórica", "Referências"]

for file in files:
  #Do we have any file that ends with a .htm or .html extension in the script's folder?
  htmlfile = re.match(r'.+(\.html|\.htm)$', file)
  if(htmlfile != None):
    xlsxData = [""] * 20
    xlsxData[2] = "Projeto"
    xlsxData[4] = 2019
    xlsxData[6] = "Não consta"
    xlsxData[7] = "Não consta"
    xlsxData[13] = "NÃO"
    print("Reading HTML file: " +file)

    with open(path + '\\' + file, 'r') as fp:            
      soup = BeautifulSoup(fp, features="html.parser", from_encoding='cp1252')              
      table = soup.find("table", class_="visualizacao")
      tbody = table.find("tbody")
      elements = tbody.find_all("tr")
      for element in elements:                
        if(len(element.contents) > 1):
          item = element.find_all("td")
          if(element.contents[1].string == "Código:"):                        
            xlsxData[0] = formatString(item[0].string)
          elif(element.contents[1].string == "Título:"):            
            xlsxData[1] = formatString(item[0].string)            
          elif(element.contents[1].string == "Categoria:"): 
            #ABRANGÊNCIA is the second data in the table cell                        
            xlsxData[3] = formatString(item[1].string)
          elif(element.contents[1].string == "Unidade Proponente:"): 
            xlsxData[5] = formatString(item[0].string)
          elif(element.contents[1].string == "Área do CNPq:"):
            #Both "Área do CNPq" and "Área principal" are contained in the same table cell     
            xlsxData[8] = formatString(item[0].string)
            xlsxData[9] = formatString(item[1].string)            
          elif(element.contents[1].string == "Público Alvo Interno:"):
            #Público alvo interno & externo            
            xlsxData[10] = formatString(item[0].string)
            xlsxData[11] = formatString(item[1].string)            
          elif(element.contents[1].string == "Fonte de Financiamento:"):                 
            xlsxData[12] = formatString(item[0].string)            
          elif(element.contents[1].string == "Coordenação:"):
            xlsxData[14] = formatString(item[0].string)            
          bold = element.find("b")
          if(bold):
            if(bold.string == "Resumo:"):
              xlsxData[15] = formatString(item[0].contents[3].string)              
            elif(bold.string == "Metodologia:"):
              xlsxData[16] = formatString(item[0].contents[3].string)              
            #elif(bold.string == "Justificativa:"):
            #  print(item[0].contents[3].string.replace('\n','').replace('\t',''))
            elif(bold.string == " Objetivos Gerais: "):              
              xlsxData[17] = formatString(item[0].contents[3].string)              
            elif(bold.string == "Fundamentação Teórica:"):              
              xlsxData[18] = formatString(item[0].contents[3].string)              
            elif(bold.string == "Referências:"):              
              xlsxData[19] = formatString(item[0].contents[3].string)              
          #endif(bold)
        #endif(len(element.contents) > 1)
      #endfor element in elements  
      for i in range(20):
        #Uncomment the 2 lines below if you do not want to alter the "Fonte de financiamento" cell in the spreadsheet
        #if(i == 12):
        #  continue        
        sheet.cell(cellIndex, i+1, xlsxData[i])
      cellIndex += 1
      print("Added contents of " +file +" into excel file")   
    #endwith open(path + '\\' + file, 'r') as fp
  #endif(htmlfile != None)
#endfor    

xlsxFile.save(path + '\\' +'PROBEX2019.xlsx') 
xlsxFile.close()
print("Saved excel file. End of script")        
            
      
       
      

    

