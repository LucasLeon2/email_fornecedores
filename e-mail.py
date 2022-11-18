import pandas as pd

dados = pd.read_excel(r'email.xlsx')

class var:
    def __init__(self, dados['Código'], dados['Alvará'], dados['Contrato social'], dados['Demonstrativo de resultado'], dados['LAO'], dados['QAF'], dados['Resp. social'], dados['ISO 9001'], dados['ISO 14001'], dados['ISO 45001']):
    self.codigo = dados['Código']
    self.alvara = dados['Alvará']
    self.contrato = dados['Contrato social']
    self.demo = dados['Demonstrativo de resultado']
    self.lao = dados['LAO']
    self.qaf = dados['QAF']
    self.termo = dados['Resp. social']
    self.iso9 = dados['ISO 9001']
    self.iso14 = dados['ISO 14001']
    self.iso45 = dados['ISO 45001']

    def filtro(self):
        for n in range(0, 100):
            print(f'Operação {n}')
            alvarat = self.alvara.isnull()
            contratot = self.contrato.isnull()
            demot = self.demo.isnull()
            laot = self.lao.isnull()
            qaft = self.qaf.isnull()
            termot = self.termo.isnull()
            iso9t = self.iso9.isnull()
            iso14t = self.iso14.isnull()
            iso45t = self.iso45.isnull()
            if alvarat[n] == 1 or laot[n] == 1 or qaft[n] == 1 or termot[n] == 1:
                print(f'E-mail deve ser enviado ao fornecedor {self.codigo[n]}')
                if contratot[n] == 1:
                    print('Contrato social também deve ser requisitado')
                if demot[n] == 1:
                    print('Demonstrativo de resultado também deve ser requisitado')
                if iso9t[n] == 1:
                    print('Norma ISO 9001 também deve ser requisitado')
                if iso14t[n] == 1:
                    print('Norma ISO 14001 também deve ser requisitado')
                if iso45t[n] == 1:
                    print('Norma ISO 45001 também deve ser requisitado')
            else:
                print(f'Nenhum e-mail deve ser enviado ao fornecedor {self.codigo[n]}')

res = var(self, dados['Código'], dados['Alvará'], dados['Contrato social'], dados['Demonstrativo de resultado'], dados['LAO'], dados['QAF'], dados['Resp. social'], dados['ISO 9001'], dados['ISO 14001'], dados['ISO 45001'])
print(res.filtro())