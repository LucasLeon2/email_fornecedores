import pandas as pd

dados = pd.read_excel(r'email.xlsx')

codigo = dados['Código']
alvara = dados['Alvará']
contrato = dados['Contrato social']
demo = dados['Demonstrativo de resultado']
lao = dados['LAO']
qaf = dados['QAF']
termo = dados['Resp. Social']
iso9 = dados['ISO 9001']
iso14 = dados['ISO 14001']
iso45 = dados['ISO 45001']

master = []
alv = []
con = []
dr = []
l = []
q = []
ter = []
i9 = []
i14 = []
i45 = []

for n in range(0,100):
    alvarat = alvara.isnull()
    contratot = contrato.isnull()
    demot = demo.isnull()
    laot = lao.isnull()
    qaft = qaf.isnull()
    termot = termo.isnull()
    iso9t = iso9.isnull()
    iso14t = iso14.isnull()
    iso45t = iso45.isnull()
    if alvarat[n] == 1 or laot[n] == 1 or qaft[n] == 1 or termot[n] == 1:
        master.append(1)
    else:
        master.append(0)
    if alvarat[n] == 1:
        alv.append('Alvará de funcionamento')
    else:
        alv.append(None)
    if laot[n] == 1:
        l.append('Licença ambiental de operação')
    else:
        l.append(None)
    if qaft[n] == 1:
        q.append('Questionário de avaliação de fornecedor(em anexo para ser preenchido e enviado de volta)')
    else:
        q.append(None)
    if termot[n] == 1:
        ter.append('Termo de responsabilidade social e ambiental(em anexo para ser assinado e enviado de volta)')
    else:
        ter.append(None)
    if contratot[n] == 1:
        con.append('Contrato social')
    else:
        con.append(None)
    if demot[n] == 1:
        dr.append('Demonstrativo de resultado de 2022')
    else:
        dr.append(None)
    if iso9t[n] == 1:
        i9.append('ISO 9001')
    else:
        i9.append(None)
    if iso14t[n] == 1:
        i14.append('ISO 14001')
    else:
        i14.append(None)
    if iso45t[n] == 1:
        i45.append('ISO 45001')
    else:
        i45.append(None)
s = 13
print(f'Fornecedor: {codigo[s]}')
print(f'Master: {master[s]}')
print(f'Alvara: {alv[s]}')
print(f'LAO: {l[s]}')
print(f'QAF: {q[s]}')
print(f'Termo: {ter[s]}')
print(f'Contrato social: {con[s]}')
print(f'Demonstrativo de resultado: {dr[s]}')
print(f'ISO 9001: {i9[s]}')
print(f'ISO 14001: {i14[s]}')
print(f'ISO 45001: {i45[s]}')
