from openpyxl import Workbook, load_workbook
import numpy as np

#Salvando a planilha com os dados em uma variável
planilhas = load_workbook('dadosTrelica.xlsx')

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
vetor_forcas = []
angulos = []
cossenos = []
senos = []
m_u = np.array([[0, 0, 0, 0]], dtype=np.float64)
m_u = m_u.T

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
        a = float(DADOS_DOS_ELEMENTOS[linha]['A'])
        k = (e*a)/comprimentos[linha]
        lista_k.append(k)
    return lista_k

#DESENVOLVIMENTO
#recebendo e armazenando os dados
lendo_armazenando (linhas_1, headers_1, DADOS_GERAIS_DO_PROBLEMA)
lendo_armazenando (linhas_2, headers_2, DADOS_DOS_NOS)
lendo_armazenando (linhas_3, headers_3, DADOS_DOS_ELEMENTOS)

#nomeando dados
qtd_nos = int(DADOS_GERAIS_DO_PROBLEMA[0]['Número de Nós'])
qtd_elementos = int(DADOS_GERAIS_DO_PROBLEMA[0]['Número de Elementos'])


#calculo do comprimento
for linha in DADOS_DOS_ELEMENTOS:
    i = int(linha['Nó I']) - 1
    j = int(linha['Nó J']) - 1
    calculo_comprimento (DADOS_DOS_NOS, i, j)

#vetor de forças
for linha in DADOS_DOS_NOS:
    px = linha['PX']
    vetor_forcas.append(px)
    py = linha['PY']
    vetor_forcas.append(py)
vetor_forcas = np.array([vetor_forcas]).T

#criando matriz global
m_global = np.zeros((qtd_nos*2 , qtd_nos*2), dtype=np.float64)

for l in range(qtd_elementos):

    i = int(DADOS_DOS_ELEMENTOS[l]['Nó I'])
    j = int(DADOS_DOS_ELEMENTOS[l]['Nó J'])
    
#matriz local
    m_local = np.array([[1 , -1], [-1 , 1]], dtype=np.float64)
    encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos)
    m_local = m_local * lista_k[l]
    
#matriz de rotacao

    R = np.array([[cossenos[l], senos[l], 0, 0],[0, 0, cossenos[l],senos[l]]], dtype=np.float64)
    R_t = R.T
    
#matriz de rigidez
    m_rigidez = np.dot(R_t, m_local)
    m_rigidez = np.dot(m_rigidez, R)

#matriz global

    m_global[2*i-2][2*i-2] = m_global[2*i-2][2*i-2] + m_rigidez[0][0]
    m_global[2*i-2][2*i-1] = m_global[2*i-2][2*i-1] + m_rigidez[0][1]
    m_global[2*i-2][2*j-2] = m_global[2*i-2][2*j-2] + m_rigidez[0][2]
    m_global[2*i-2][2*j-1] = m_global[2*i-2][2*j-1] + m_rigidez[0][3]
    
    m_global[2*i-1][2*i-2] = m_global[2*i-1][2*i-2] + m_rigidez[1][0]
    m_global[2*i-1][2*i-1] = m_global[2*i-1][2*i-1] + m_rigidez[1][1]
    m_global[2*i-1][2*j-2] = m_global[2*i-1][2*j-2] + m_rigidez[1][2]
    m_global[2*i-1][2*j-1] = m_global[2*i-1][2*j-1] + m_rigidez[1][3]

    m_global[2*j-2][2*i-2] = m_global[2*j-2][2*i-2] + m_rigidez[2][0]
    m_global[2*j-2][2*i-1] = m_global[2*j-2][2*i-1] + m_rigidez[2][1]
    m_global[2*j-2][2*j-2] = m_global[2*j-2][2*j-2] + m_rigidez[2][2]
    m_global[2*j-2][2*j-1] = m_global[2*j-2][2*j-1] + m_rigidez[2][3]

    m_global[2*j-1][2*i-2] = m_global[2*j-1][2*i-2] + m_rigidez[3][0]
    m_global[2*j-1][2*i-1] = m_global[2*j-1][2*i-1] + m_rigidez[3][1]
    m_global[2*j-1][2*j-2] = m_global[2*j-1][2*j-2] + m_rigidez[3][2]
    m_global[2*j-1][2*j-1] = m_global[2*j-1][2*j-1] + m_rigidez[3][3]


#restricoes
m_global_1 = m_global.copy()
v_forcas_1 = vetor_forcas.copy()

for i in range(qtd_nos):
    acm = (i*2) + 1 
    if DADOS_DOS_NOS[i]['DX']:
        vetor_forcas[i*2] = 0
        for linha in range(qtd_nos*2):
            m_global[linha][i] = 0
        for coluna in range(qtd_nos*2):
            m_global[i][coluna] = 0
        m_global[i*2][i*2] = 1

    if DADOS_DOS_NOS[i]['DY']:
        vetor_forcas[acm] = 0
        for linha in range(qtd_nos*2):
            m_global[linha][acm] = 0
        for coluna in range(qtd_nos*2):
            m_global[acm][coluna] = 0
        m_global[acm][acm] = 1


result_U = np.linalg.solve(m_global, vetor_forcas)
print('DESLOCAMENTOS')
print(result_U)

reacao = np.dot(m_global_1, result_U)
reacao1 = reacao - v_forcas_1

print('\nREAÇÕES')

for l in range (qtd_nos):

    i = int(DADOS_DOS_NOS[l]['x'])
    j = int(DADOS_DOS_NOS[l]['y'])

    fx = reacao1[2*l]
    fy = reacao1[2*l + 1]

    if DADOS_DOS_NOS[l]['DX'] == 1:
        print (f'X: {i}, Y: {j} - Nó {l}: {fx}')

    if DADOS_DOS_NOS[l]['DY'] == 1:
        print (f'X: {i}, Y: {j} - Nó {l}: {fy}')


print('\nESFORÇOS')
for l in range (qtd_elementos):

    i = int(DADOS_DOS_ELEMENTOS[l]['Nó I'])
    j = int(DADOS_DOS_ELEMENTOS[l]['Nó J'])

    m_u[0] = result_U [2*i-2]
    m_u[1] = result_U [2*i-1]
    m_u[2] = result_U [2*j-2]
    m_u[3] = result_U [2*j-1]

    R = np.array([[cossenos[l], senos[l], 0, 0],[0, 0, cossenos[l],senos[l]]], dtype=np.float64)
    m_local = np.array([[1 , -1], [-1 , 1]], dtype=np.float64)
    encontrar_K (DADOS_DOS_ELEMENTOS, comprimentos, lista_k, qtd_elementos)
    m_local = m_local * lista_k[l]

    esforcos = np.dot (R, m_u)

    esforcos = np.dot (m_local, esforcos)

    print (f"Elemento {l+1}: {esforcos}")