import win32com.client as win32
from tkinter import Entry
import os
from tkinter import*

def programa():
    #PEGAR VALORES DA TELA:
    gmailTela = str(gmailDe.get())
    assuntoTela = str(assunto.get())
    #COMO A MENSAGEMM É DO TIPO TEXTO, PRECISSAÇE LER DO INICIO DA 1° LINHA ATÉ A ULTIMA: "1.0", "end-1c"
    mensagemTela = str(mensagem.get("1.0", 'end-1c'))
    quantVezes = int(vezes.get())

    #PASSAR VALORES PARA A EXECUÇÃO
    cont = 0
    while cont < quantVezes:
        cont +=1

        #CRIAR INTEGRAÇÃO COM OUTLOOK
        outlook = win32.Dispatch('outlook.application')

        #CRIAR UM EMAIL
        email = outlook.CreateItem(0)

        #EMAIL 1°
        email.To = f"{gmailTela}"
        email.Subject = f"{assuntoTela}"
        email.HTMLBody = f"""
        <p>{mensagemTela}</p>
        """
        email.Send()

        #MOSTRAR NA TELA
        res = f"{quantVezes}° EMAIL ENVIADO COM SUCESSO!!!"
        resul['text'] = res


#TELA_____________________________________________________
tela = Tk()
tela.resizable(0, 0)
tela.title("GMAIL OUT-LOOK")
tela.geometry("470x660")

texto_orientacao = Label(tela, text="GMAIL VIA OUT-LOOK", foreground = "blue")
texto_orientacao.grid(column = 0, row=0, padx=0, pady=20)
texto_orientacao.config(font=("Helvitica, 18"))

#PARA QUEM ENVIAR O GMAIL
Label(tela, text="PARA: ", foreground="black", anchor=W).place(x=10, y=80, width=100, height=20)
gmailDe = Entry(tela)
gmailDe.place(x=10, y=100, width=400, height=30)
gmailDe.config(font=("Helvitica, 15"))

#ASSUNTO:
Label(tela, text="ASSUNTO: ", foreground="black", anchor=W).place(x=10, y=150, width=100, height=20)
assunto = Entry(tela)
assunto.place(x=10, y=170, width=300, height=30)
assunto.config(font=("Helvitica, 15"))

#MENSAGEM:
Label(tela, text="MENSAGEM:: ", foreground="black", anchor=W).place(x=10, y=220, width=100, height=20)
mensagem = Text(tela)
mensagem.place(x=10, y=240, width=400, height=250)
mensagem.config(font=("Helvitica, 15"))

#QUANTIDADE DE VEZES
Label(tela, text="QUANTIDADE DE VEZES: ", foreground="black", anchor=W).place(x=10, y=510, width=200, height=20)
vezes = Entry(tela)
vezes.place(x=10, y=530, width=100, height=30)
vezes.config(font=("Helvitica, 15"))

#BOTÃO-ENVIAR:
botao = Button(tela, text="ENVIAR", command = programa, foreground = "red", anchor = W).place(x=10, y=590, width=60, height=40)

#RESULTADO
resul = Label(tela, text = '', foreground = "green")
resul.place(x=10, y=640, width=400, height=25)
resul.config(font = ("Helvitica, 13"))

tela.mainloop()









