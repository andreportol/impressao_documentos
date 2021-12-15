import win32print  # impressora
import win32api
import os
from openpyxl import load_workbook


class App2:
    def abrir_planilha_controle(self):
        self.caminho1 = r"C:\Users\andre.porto\Desktop\teste"
        self.caminho2 = r"C:\Users\andre.porto\Desktop\teste\pdf"
        # abrir planilha controle1 do excel
        planilha_nova = load_workbook(self.caminho1 + r'\controle.xlsx')
        self.celula = planilha_nova.active

    def configurar_impressora(self):
        # 2 é o valor padrão
        lista_impressoras = win32print.EnumPrinters(2)
        # exibe a lista de impressoras instaladas
        # print (lista_impressoras)
        # impressora n. 4
        # para escolha da impressora,
        # contagem de trás para frente
        impressora = lista_impressoras[4]
        # 2 É A CONFIGURAÇÃO PADRÃO
        win32print.SetDefaultPrinter(impressora[2])

    def imprimir_documentos(self):
        # criando uma lista com os arquivos obtidos dentro da pasta 'pdf'
        lista_de_arquivos = os.listdir(self.caminho2)
        c = 2  # iniciar na posição 2 da célula C
        b = 2  # iniciar na posição 2 da célula D
        for arquivo in lista_de_arquivos:
            if(int(self.celula[f'C{c}'].value) % 2 != 0):
                for _ in range(self.celula[f'B{b}'].value):
                    win32api.ShellExecute(
                        0, "print", arquivo, None, self.caminho2, 0)
            else:
                print(f'Arquivo:{arquivo} -> par')
                win32api.ShellExecute(0, "open", arquivo,
                                      None, self.caminho2, 0)
            c += 1
            b += 1


if __name__ == '__main__':
    aplicativo = App2()
    aplicativo.abrir_planilha_controle()
    aplicativo.configurar_impressora()
    aplicativo.imprimir_documentos()
    print(' FIM -> PROGRAMA FINALIZADO....AGUARDE IMPRESSÃO!!')
