import customtkinter 
from pegar_moedas import nomes_moedas, conversoes_disponiveis  
from pegar_cotacao import pegar_cotacao_moeda   


customtkinter.set_appearance_mode("blue")
customtkinter.set_default_color_theme("dark-blue")

janela=customtkinter.CTk()
janela.title("Sistema de Conversão de Moedas")
janela.geometry("1000x1000")


label_inferior = customtkinter.CTkLabel(janela, text=" Maycon Depolli dos Santos ",bg_color="lightgrey",font=("",20))

# Criação do label (texto)


dic_conversoes_disponiveis=conversoes_disponiveis()
#dic_conversoes_disponiveis["USD"]=> ["BRL,"CAD","BTC","AUD"]

titulo=customtkinter.CTkLabel(janela, text="Converta a Moeda",text_color="green" ,font=("",20))
texto_moeda_origem=customtkinter.CTkLabel(janela,text="Selecione a Moeda de origem:",text_color="blue",font=("",15))
texto_moeda_destino=customtkinter.CTkLabel(janela,text="Selecione a Moeda de destino:",text_color="black",font=("",15))

def carregar_moedas_destino(moeda_selecionada):
    lista_moedas_destino=dic_conversoes_disponiveis[moeda_selecionada]
    campo_moeda_destino.configure(values=lista_moedas_destino)
    campo_moeda_destino.set(lista_moedas_destino[0])
campo_moeda_origem = customtkinter.CTkOptionMenu(janela,values=list(dic_conversoes_disponiveis.keys()),
                                                command=carregar_moedas_destino)                                   
campo_moeda_destino = customtkinter.CTkOptionMenu(janela,values=["Selecione uma moeda de origem"])

def converter_moeda():
    moeda_origem=campo_moeda_origem.get()
    moeda_destino=campo_moeda_destino.get()
    if moeda_origem and moeda_destino:
        cotacao=pegar_cotacao_moeda(moeda_origem,moeda_destino)
        texto_cotacao_moeda.configure(text=f"1 {moeda_origem}={cotacao}{moeda_destino}")
botao_converter=customtkinter.CTkButton(janela,text="Converter",fg_color="black",command=converter_moeda)

lista_moedas=customtkinter.CTkScrollableFrame(janela)
texto_cotacao_moeda=customtkinter.CTkLabel(janela,text="")
moedas_disponiveis= nomes_moedas()
for codigo_moeda in moedas_disponiveis:
    nome_moeda=moedas_disponiveis[codigo_moeda]
    texto_moeda=customtkinter.CTkLabel(lista_moedas, text=f"{codigo_moeda}:{nome_moeda}")
    texto_moeda.pack()

moeda1=customtkinter.CTkLabel(lista_moedas,text="USD: Dólar americano")
moeda2=customtkinter.CTkLabel(lista_moedas,text="BRL: Real Brasileiro")
moeda3=customtkinter.CTkLabel(lista_moedas,text="BTC: Bitcoin")
moeda4=customtkinter.CTkLabel(lista_moedas,text="EUR: Euro")
moeda1.pack();
moeda2.pack();
moeda3.pack();
moeda4.pack();
titulo.pack(padx=10,pady=10)
texto_moeda_origem.pack(padx=10,pady=10)
campo_moeda_origem.pack(padx=10)
texto_moeda_destino.pack(padx=10,pady=10)
campo_moeda_destino.pack(padx=10)
botao_converter.pack(padx=10,pady=10)
texto_cotacao_moeda.pack(padx=10,pady=10)
lista_moedas.pack(padx=10,pady=10)
label_inferior.pack(side=customtkinter.BOTTOM, fill=customtkinter.X, pady=10)
janela.mainloop()