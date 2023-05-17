from openpyxl import Workbook, load_workbook
import numpy as np

#Salvando a planilha com os dados em uma variável
planilhas = load_workbook('dadosMEF1.xlsx')

#Selecionando e salvando a plan(sheet) "Dados dos Elementos"  
sheet_1 = planilhas['DADOS GERAIS DO PROBLEMA']
sheet_2 = planilhas['DADOS DOS NÓS']
sheet_3 = planilhas['DADOS DOS ELEMENTOS']

#pegar todas as linhas da planilha (propriedade rows)
#propriedade rows retorna um objeto gerador
linhas_1 = sheet_1.rows
linhas_2 = sheet_2.rows
linhas_3 = sheet_3.rows

#lendo as linhas
headers_1 = next(linhas_1)
headers_2 = next(linhas_2)
headers_3 = next(linhas_3)

#celula.value -> valor da celular
headers_1 = [celula.value for celula in next(linhas_1)]
headers_2 = [celula.value for celula in next(linhas_2)]
headers_3 = [celula.value for celula in next(linhas_3)]

#declaração de variaveis
DADOS_GERAIS_DO_PROBLEMA = []
DADOS_DOS_NOS = []
DADOS_DOS_ELEMENTOS = []
comprimentos = []
lista_k = []
m_local =[]
m_global = []
vetor_forças_distribuida_1 = []
vetor_forças_distribuida_2 = []
vetor_carga_pontual_x = []
vetor_carga_pontual_y = []

#funções 
def lendo_armazenando (linhas, headers, lista_linhas):
    for linha in linhas:
        dados = {}
        for titulo,celula in zip(headers, linha):
            dados[titulo] = celula.value
    
        lista_linhas.append(dados)

#calculo do comprimento a partir das coordenadas e adiciona em uma lista
def calculo_comprimento (DADOS_DOS_NOS, i, j):
    x_1 = DADOS_DOS_NOS[i]['x']
    y_1 = DADOS_DOS_NOS[i]['y']
    x_2 = DADOS_DOS_NOS[j]['x']
    y_2 = DADOS_DOS_NOS[j]['y']
    x = x_2 - x_1
    y = y_2 - y_1 
    L = ((x**2)+(y**2))**(1/2)
    comprimentos.append(L)

def encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos):
    for linha in range(qtd_elementos):
        e = float(DADOS_DOS_ELEMENTOS[linha]['E'])
        a = float(DADOS_DOS_ELEMENTOS[linha]['A'])
        k = (e*a)/comprimentos[linha]
        lista_k.append(k)
    return lista_k

#des
lendo_armazenando (linhas_1, headers_1, DADOS_GERAIS_DO_PROBLEMA)
lendo_armazenando (linhas_2, headers_2, DADOS_DOS_NOS)
lendo_armazenando (linhas_3, headers_3, DADOS_DOS_ELEMENTOS)

for linha in DADOS_DOS_ELEMENTOS:
    i = int(linha['Nó I']) - 1
    j = int(linha['Nó J']) - 1
    calculo_comprimento (DADOS_DOS_NOS, i, j)

qtd_nos = len(DADOS_DOS_NOS)
qtd_elementos = len(DADOS_DOS_ELEMENTOS)   
encontrar_K(DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos)
m_global = np.zeros((qtd_elementos+1 , qtd_elementos+1), dtype=np.float64)

for linha in range(qtd_elementos):
    l = comprimentos[linha]
    if (linha > 0):
        qx_i = DADOS_DOS_ELEMENTOS[linha]['Qx(I)']
        fd_1 = (qx_i*l)/2
        qx_f = DADOS_DOS_ELEMENTOS[linha-1]['Qx(J)']
        fd_2 = (qx_f*l)/2
        fdx = fd_1 + fd_2
        vetor_forças_distribuida_1.append(fdx)
        qy_i = DADOS_DOS_ELEMENTOS[linha]['Qy(I)']
        fd_3 = (qy_i*l)/2
        qy_f = DADOS_DOS_ELEMENTOS[linha-1]['Qy(J)']
        fd_4 = (qy_f*l)/2
        fdy = fd_3 + fd_4
        vetor_forças_distribuida_2.append(fdy)
    else:
        qx_i = DADOS_DOS_ELEMENTOS[linha]['Qx(I)']
        qy_i = DADOS_DOS_ELEMENTOS[linha]['Qy(I)']
        fd_5 = (qx_i*l)/2
        fd_6 = (qy_i*l)/2
        vetor_forças_distribuida_1.append(fd_5)
        vetor_forças_distribuida_2.append(fd_6)

fd_7 = (qx_i*l)/2
fd_8 = (qy_i*l)/2
vetor_forças_distribuida_1.append(fd_7)
vetor_forças_distribuida_2.append(fd_8)

for linha in DADOS_DOS_NOS:
    px = linha['PX']
    py = linha['PY']
    vetor_carga_pontual_x.append(px)
    vetor_carga_pontual_y.append(py)

matriz_forças_distribuida_1 = np.array([vetor_forças_distribuida_1]).T
matriz_forças_distribuida_2 = np.array([vetor_forças_distribuida_2]).T
matriz_carga_pontual_x = np.array([vetor_carga_pontual_x]).T
matriz_carga_pontual_y = np.array([vetor_carga_pontual_y]).T


vetor_forças = matriz_forças_distribuida_1 + matriz_forças_distribuida_2 + matriz_carga_pontual_x + matriz_carga_pontual_y

for linha in range(qtd_elementos):
    
    m_local = np.array([[1 , -1], [-1 , 1]], dtype=np.float64)
    m_local = m_local * lista_k[linha]
    i = int(DADOS_DOS_ELEMENTOS[linha]['Nó I'])
    j = int(DADOS_DOS_ELEMENTOS[linha]['Nó J'])
    m_global[i-1][i-1] = m_global[i-1][i-1] + m_local[0][0]
    m_global[i-1][i] = m_global[i-1][i] + m_local[0][1]
    m_global[j-1][j-2] = m_global[j-1][j-2] + m_local[1][0]
    m_global[j-1][j-1] = m_global[j-1][j-1] + m_local[1][1]

for i in range(qtd_nos):
    if DADOS_DOS_NOS[i]['DX']:
        vetor_forças[i] = 0
        for linha in range(qtd_nos):
            m_global[i][linha] = 0
        for coluna in range(qtd_nos):
            m_global[coluna][i] = 0
        m_global[i][i] = 1

result_U = np.linalg.solve(m_global, vetor_forças)

print (result_U)