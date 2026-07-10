import xmltodict

def nomes_moedas ():
    with open (r"C://Users//msantos//Downloads//Projeto//moedas.xml","rb") as arquivo_moedas:
        dic_moedas = xmltodict.parse(arquivo_moedas)

    moedas = dic_moedas ["xml"]
    return moedas

def conversoes_disponiveis():

    with open (r"C://Users//msantos//Downloads//Projeto//conversoesmoedas.xml","rb") as arquivo_conversoesmoedas:
        dic_conversoesmoedas=xmltodict.parse(arquivo_conversoesmoedas)

    conversoesmoedas= dic_conversoesmoedas["xml"]
    dic_conversoes_disponiveis={}
    for par_conversao in conversoesmoedas:
        moeda_origem,moeda_destino=par_conversao.split("-")
        if moeda_origem in dic_conversoes_disponiveis:
            dic_conversoes_disponiveis[moeda_origem].append (moeda_destino)
        else:
            dic_conversoes_disponiveis[moeda_origem] = [moeda_destino] 
    return dic_conversoes_disponiveis   
    {"USD": ["BRL","BTC","CAD","AUD","AED"]}

conversoes_disponiveis()