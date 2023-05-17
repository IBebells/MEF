from openpyxl import Workbook, load_workbook
import numpy as np

#Salvando a planilha com os dados em uma variável
planilhas = load_workbook('dadosVIGA1.xlsx')

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
cossenos = []
senos = []
v_forcas_pontual = []

#FUNCOES
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
    cos = x/L
    sen = y/L
    comprimentos.append(L)
    cossenos.append(cos)
    senos.append(sen)

def encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos):
    for linha in range(qtd_elementos):
        e = float(DADOS_DOS_ELEMENTOS[linha]['E'])
        a = float(DADOS_DOS_ELEMENTOS[linha]['IZ'])
        k = (e*a)/comprimentos[linha]
        lista_k.append(k)
    return lista_k

#DESENVOLVIMENTO
#recebendo e armazenando os dados
lendo_armazenando (linhas_1, headers_1, DADOS_GERAIS_DO_PROBLEMA)
lendo_armazenando (linhas_2, headers_2, DADOS_DOS_NOS)
lendo_armazenando (linhas_3, headers_3, DADOS_DOS_ELEMENTOS)

#nomeando dados
qtd_nos = (int(DADOS_GERAIS_DO_PROBLEMA[0]['Número de nós']))
qtd_elementos = (int(DADOS_GERAIS_DO_PROBLEMA[0]['Número de Elementos']))

v_forcas_global = np.zeros((qtd_nos*2 , 1))

#calculo do comprimento
for linha in DADOS_DOS_ELEMENTOS:
    i = int(linha['Nó I']) - 1
    j = int(linha['Nó J']) - 1
    calculo_comprimento (DADOS_DOS_NOS, i, j)

#vetor de carregamento pontual
for linha in DADOS_DOS_NOS:
    py = linha['PY']
    v_forcas_pontual.append(py)
    mz = linha['MZ']
    v_forcas_pontual.append(mz)
v_forcas_pontual = np.array([v_forcas_pontual]).T

m_global = np.zeros((qtd_nos*2 , qtd_nos*2), dtype=np.float64)

for l in range(qtd_elementos):

    i = int(DADOS_DOS_ELEMENTOS[l]['Nó I'])
    j = int(DADOS_DOS_ELEMENTOS[l]['Nó J'])
    qy_i = DADOS_DOS_ELEMENTOS[l]['Qy(I)']
    qy_j = DADOS_DOS_ELEMENTOS[l]['Qy(J)']

#matriz local
    L = comprimentos[l]

    m_local = np.array([[12/L**2, 6/L, -12/L**2, 6/L], [6/L, 4, -6/L , 2], [-12/L**2 , -6/L, 12/L**2, -6/L], [6/L , 2, -6/L, 4]], dtype=np.float64)
    encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos)
    m_local = m_local * lista_k[l]

#vetor de forcas
    v_forcas_distribuida = np.array([((7/20)*L*qy_i)+((3/20)*L*qy_j), ((1/20)*(L**2)*qy_i)+((1/30)*(L**2)*qy_j), ((3/20)*L*qy_i)+((7/20)*L*qy_j), ((-1/30)*(L**2)*qy_i)+((-1/20)*(L**2)*qy_j)], dtype=np.float64)
    v_forcas_distribuida = np.array([v_forcas_distribuida]).T 

#matriz de rigidez global    
    m_global[2*i-2][2*i-2] = m_global[2*i-2][2*i-2] + m_local[0][0]
    m_global[2*i-2][2*i-1] = m_global[2*i-2][2*i-1] + m_local[0][1]
    m_global[2*i-2][2*j-2] = m_global[2*i-2][2*j-2] + m_local[0][2]
    m_global[2*i-2][2*j-1] = m_global[2*i-2][2*j-1] + m_local[0][3]
    
    m_global[2*i-1][2*i-2] = m_global[2*i-1][2*i-2] + m_local[1][0]
    m_global[2*i-1][2*i-1] = m_global[2*i-1][2*i-1] + m_local[1][1]
    m_global[2*i-1][2*j-2] = m_global[2*i-1][2*j-2] + m_local[1][2]
    m_global[2*i-1][2*j-1] = m_global[2*i-1][2*j-1] + m_local[1][3]

    m_global[2*j-2][2*i-2] = m_global[2*j-2][2*i-2] + m_local[2][0]
    m_global[2*j-2][2*i-1] = m_global[2*j-2][2*i-1] + m_local[2][1]
    m_global[2*j-2][2*j-2] = m_global[2*j-2][2*j-2] + m_local[2][2]
    m_global[2*j-2][2*j-1] = m_global[2*j-2][2*j-1] + m_local[2][3]

    m_global[2*j-1][2*i-2] = m_global[2*j-1][2*i-2] + m_local[3][0]
    m_global[2*j-1][2*i-1] = m_global[2*j-1][2*i-1] + m_local[3][1]
    m_global[2*j-1][2*j-2] = m_global[2*j-1][2*j-2] + m_local[3][2]
    m_global[2*j-1][2*j-1] = m_global[2*j-1][2*j-1] + m_local[3][3]

    v_forcas_global[2*i-2] = v_forcas_global[2*i-2] + v_forcas_distribuida[0]
    v_forcas_global[2*i-1] = v_forcas_global[2*i-1] + v_forcas_distribuida[1]
    v_forcas_global[2*j-2] = v_forcas_global[2*j-2] + v_forcas_distribuida[2]
    v_forcas_global[2*j-1] = v_forcas_global[2*j-1] + v_forcas_distribuida[3]

v_forcas_global = v_forcas_global + v_forcas_pontual

m_global_s_rest = m_global.copy()
v_forcas_s_rest = v_forcas_global.copy()

#condicoes de contorno
for i in range(qtd_nos):
    acm = (i*2) + 1

    if DADOS_DOS_NOS[i]['DY']:
        v_forcas_global[i*2] = 0
        for linha in range(qtd_nos*2):
            m_global[linha][i*2] = 0
        for coluna in range(qtd_nos*2):
            m_global[i*2][coluna] = 0
        m_global[i*2][i*2] = 1

    if DADOS_DOS_NOS[i]['RZ']:
        v_forcas_global[acm] = 0
        for linha in range(qtd_nos*2):
            m_global[linha][acm] = 0
        for coluna in range(qtd_nos*2):
            m_global[acm][coluna] = 0
        m_global[acm][acm] = 1

result_U = np.linalg.solve(m_global, v_forcas_global)
print('DEFORMAÇÕES')
print(result_U)

reacao = np.dot(m_global_s_rest, result_U)
reacao1 = reacao - v_forcas_s_rest
print('REAÇÕES')
print(reacao1)

m_u = np.array([[0, 0, 0, 0]], dtype=np.float64)
m_u = m_u.T

for l in range (qtd_elementos):

    i = int(DADOS_DOS_ELEMENTOS[l]['Nó I'])
    j = int(DADOS_DOS_ELEMENTOS[l]['Nó J'])
    qy_i = DADOS_DOS_ELEMENTOS[l]['Qy(I)']
    qy_j = DADOS_DOS_ELEMENTOS[l]['Qy(J)']

    m_u[0] = result_U [2*i-2]
    m_u[1] = result_U [2*i-1]
    m_u[2] = result_U [2*j-2]
    m_u[3] = result_U [2*j-1]

    #matriz local
    L = comprimentos[l]
    m_local = np.array([[12/L**2, 6/L, -12/L**2, 6/L], [6/L, 4, -6/L , 2], [-12/L**2 , -6/L, 12/L**2, -6/L], [6/L , 2, -6/L, 4]], dtype=np.float64)
    encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos)
    m_local = m_local * lista_k[l]

    v_forcas_distribuida = np.array([((7/20)*L*qy_i)+((3/20)*L*qy_j), ((1/20)*(L**2)*qy_i)+((1/30)*(L**2)*qy_j), ((3/20)*L*qy_i)+((7/20)*L*qy_j), ((-1/30)*(L**2)*qy_i)+((-1/20)*(L**2)*qy_j)], dtype=np.float64)
    v_forcas_distribuida = np.array([v_forcas_distribuida]).T 

    esforcos = np.dot (m_local, m_u)
    esforcos = esforcos - v_forcas_distribuida
    print (f"Esforço da Elemento {l+1}: {esforcos}")