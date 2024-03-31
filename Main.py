import tkinter as tk
from tkinter import ttk
from threading import Thread
import cv2
from pyzbar import pyzbar
import pandas as pd

pdProdutos = pd.read_excel('Produtos.xlsx', dtype={'Numero do Código de Barras': str})

pdProdutos['Numero do Código de Barras'] = pdProdutos['Numero do Código de Barras']

codigos_detectados = set(pdProdutos['Numero do Código de Barras'])

class YourClassName: 
    def __init__(self):
        self.pdProdutos = pd.read_excel('Produtos.xlsx')
        self.is_running = True
    def a(self, TelaS):

        self.TelaS = TelaS
        self.TelaS.title("Scanner")
        self.TelaS.geometry("433x339")
        self.TelaS.configure(bg='#03246D')

        caminho_icone = "ICON.ico"
        self.TelaS.iconbitmap(caminho_icone)

        self.btnSalvar = tk.Button(self.TelaS, text="Salvar", command=self.dowload_dos_dados, font=('Ebrima', 12, 'bold'), bg='#010F2E', fg='white')
        self.btnSalvar.place(relx=0.27, rely=0.74, relwidth=0.44, relheight=0.15)

        self.labeltitulo = ttk.Label(self.TelaS, text="Solis Scanner", background='#03246D', font=("Ebrima", 14, "bold"), foreground='white')
        self.labeltitulo.place(relx=0.35, rely=0.06, relwidth=0.68, relheight=0.12)

        self.txtNome = ttk.Label(self.TelaS, text="Nome do Produto:", background='#03246D', font=("Ebrima", 12, "bold"), foreground='white')
        self.txtNome.place(relx=0.09, rely=0.25, relwidth=0.43, relheight=0.09)

        self.txtPreco = ttk.Label(self.TelaS, text="Preço do Produto:", background='#03246D', font=("Ebrima", 12, "bold"), foreground='white')
        self.txtPreco.place(relx=0.09, rely=0.48, relwidth=0.57, relheight=0.09)

        self.entryNome = ttk.Entry(self.TelaS, font=("Ebrima", 12, "bold"))
        self.entryNome.place(relx=0.70, rely=0.30, relwidth=0.5, relheight=0.09, anchor='center')

        self.entryPreco = ttk.Entry(self.TelaS, font=("Ebrima", 12, "bold"))
        self.entryPreco.place(relx=0.45, rely=0.48, relwidth=0.5, relheight=0.09)

        self.TelaS.mainloop()

    def dowload_dos_dados(self):
        codigos_detectados.add(barcode_value)
        nome_produto = self.entryNome.get()
        preco_produto = self.entryPreco.get()
        novo_produto = {'Numero do Código de Barras': barcode_value, 'Nome do Produto': nome_produto, 'Preço': preco_produto}
        if not self.pdProdutos.empty:
            self.pdProdutos = pd.concat([self.pdProdutos, pd.DataFrame([novo_produto])], ignore_index=True)
        else:
            self.pdProdutos = pd.DataFrame([novo_produto])
        self.TelaS.destroy()


cap = cv2.VideoCapture(0)
while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break
    barcodes = pyzbar.decode(frame)

    for barcode in barcodes:
        (x, y, w, h) = barcode.rect
        cv2.rectangle(frame, (x, y), (x + w, y + h), (255,255,255), 2)
        barcode_value = barcode.data.decode("utf-8")
        if str(barcode_value) in codigos_detectados:
            status = "Ja cadastrado"
        else:
            status = barcode_value
        cv2.putText(frame, status, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255,255,255), 2)

        if cv2.waitKey(1) & 0xFF == ord('a'):
            if barcode_value not in codigos_detectados:
                your_instance = YourClassName()
                your_instance.a(tk.Tk())


    cv2.imshow("Solis Scanner", frame)
    if cv2.waitKey(1) & 0xFF == ord('s'):
        break

cap.release()
cv2.destroyAllWindows()
your_instance.pdProdutos.to_excel('Produtos.xlsx', index=False)