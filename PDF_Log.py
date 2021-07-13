from PDF_UI import *
import fitz
import os
import openpyxl
from PyQt5.QtWidgets import QFileDialog
class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, *args, **kwargs):
        QtWidgets.QMainWindow.__init__(self, *args, **kwargs)
        self.setupUi(self)
        self.btnSalir.clicked.connect(self.close)
        self.btnAbrir.clicked.connect(self.funci)
    def agregaDatos(self,a,name,tipo):
        book = openpyxl.load_workbook('ejercicios528.xlsx',data_only=False)
        hoja = book.active
        cont=1
        colu=[]
        col=[]
        if tipo == '11':
            for i in range(131,240):
            #contanto los primeros 10datos
                if cont <11:
                    colu.append(a[i])
                    cont+=1
                else:
                    col.append(colu)
                    colu=[]
                    colu.append(a[i])
                    cont=2
            cont =1
            colu=[]
            for i in range(11,111):
            #contanto los primeros 10datos
                if cont <11:
                    colu.append(a[i])
                    cont+=1
                else:
                    col.append(colu)
                    colu=[]
                    colu.append(a[i])
                    cont=2
            col.append(colu)
        else:
            for i in range(11,111):
            #contanto los primeros 10datos
                if cont <11:
                    colu.append(a[i])
                    cont+=1
                else:
                    col.append(colu)
                    colu=[]
                    colu.append(a[i])
                    cont=2
            col.append(colu)
            cont =1
            colu=[]
            for i in range(131,240):
            #contanto los primeros 10datos
                if cont <11:
                    colu.append(a[i])
                    cont+=1
                else:
                    col.append(colu)
                    colu=[]
                    colu.append(a[i])
                    cont=2
            
        for j in range(0,20):
            for i in range(0,10):
                hoja.cell(row = 2+j,column =4+i, value=int(col[j][i]))
        book.save(name+'.xlsx')
            

    def funci(self):
        pdf = QFileDialog.getOpenFileName(self, 'Open a file', '','All Files (*.*)')
        docum = fitz.open(pdf[0])
        np=docum.pageCount
        iniciopagina=[]
        for i in range(1,np+1):
            iniciopagina.append((247*i)-1)
        salida = open("prueba.txt","wb")
        for pag in docum:
            texto =pag.getText().encode("utf8")
            salida.write(texto)
        salida.close()
        txt = open("prueba.txt","r",encoding="utf8")
        cont= txt.readlines()
        txt.close()
        noneElement=[]
        for i in range(0,len(cont)):
            if cont[i] == " \n":
                noneElement.append(i)
        a = len(noneElement)
        for i in range(0,a):
            cont.remove(" \n")
        cont2=[]
        for i in range(0,len(cont)):
            aux=cont[i].split(sep=' \n')
            cont2.append(aux[0])
        pagi=[]
        pagaux=[]
        j=0
        for i in range(0,len(cont2)):
            pagaux.append(cont2[i])
            if i in iniciopagina:
                pagi.append(pagaux)
                pagaux=[]
        for i in range(0,len(pagi)):
            self.agregaDatos(pagi[i],'Prueba'+str(i),pagi[i][1])
        

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()