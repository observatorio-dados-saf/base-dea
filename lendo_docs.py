from re import search, findall, IGNORECASE
from caminho import pasta #caminho da pasta com os relatórios DEA
import win32com.client
import docx2txt
import pandas as pd
import os

relatorios  = os.listdir(pasta)#lista com os nomes dos relatórios
base_relatorios = pd.DataFrame()
padrao_mes = r'\b(?:janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b'



def texto_word(arquivo):#arquivo = caminho do arquivo --- função para ler os arquivos .doc
    word_app = win32com.client.Dispatch("Word.Application")
    doc = word_app.Documents.Open(arquivo)
    texto = doc.Content.Text
    doc.Close()
    word_app.Quit()
    return texto

def base_word(conteudo):#função para obter informações do relatório e transformar em um df
  
    empenho = findall('Empenho nº (.*?),', conteudo)#todo rap possui empenho

    credor = search('para (.*?),', conteudo)[1]
    proc_e_cont = findall(r' (\d{1,7}/\d{2,4})', conteudo)
    processo = proc_e_cont[0]
    contrato = proc_e_cont[1]
    competencia = findall(padrao_mes, conteudo, IGNORECASE)[0]
    valores = findall('R\$(.*?)\(', texto)#valores
    valor_dea = valores[-1]
  
    if len(empenho) == 1:
        n_empenho = empenho[0]# nota de empenho caso RAP
        valor_rap = valores[1]
    else:
        n_empenho = ''
        valor_rap = ''
  
    dic_dados = {'PROCESSO':[processo], 'CREDOR':[credor],'N CONTRATO': [contrato], 'MES COMPETENCIA':[competencia.replace('.','')],
                'N EMPENHO': [n_empenho],'RAP':[valor_rap.replace('.','')],'DEA':[valor_dea.replace('.','')]}
  
    dados_relatorio = pd.DataFrame(dic_dados)
    return dados_relatorio



for relatorio in relatorios:
    try:
        if '.doc' in relatorio:
            texto = texto_word(r'{}\{}'.format(pasta, relatorio))
        if '.docx' in relatorio:
            texto = docx2txt.process(r'{}\{}'.format(pasta, relatorio))
        
        valores = findall('R\$ (\d+\.\d+\,\d+)', texto)#valores
        base_relatorios = pd.concat([base_relatorios, base_word(texto)], ignore_index=True)
    
    except:
        dic_erro = {'PROCESSO':[relatorio]}
        dados_erro = pd.DataFrame(dic_erro)
        base_relatorios = pd.concat([base_relatorios, dados_erro], ignore_index=True)

#removendo espaços e trocando caracters
base_relatorios['RAP'] = base_relatorios['RAP'].replace({'R$':'',' ':'', ',':'.'}, regex = True)
base_relatorios['DEA'] = base_relatorios['DEA'].replace({'R$':'',' ':'', ',':'.'}, regex = True)


#exportando para excel
base_relatorios.to_excel(r'base_relatorios_dea.xlsx', index=False)