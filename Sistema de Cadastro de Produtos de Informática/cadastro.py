import os
import mimetypes
import tkinter as tk
from tkinter import filedialog
from tkinter import PhotoImage
from tkinter import messagebox
from supabase import create_client, Client

from dotenv import load_dotenv
import uuid 

load_dotenv()

url = os.getenv("SUPABASE_URL")
key = os.getenv("SUPABASE_KEY")

if not url or not key:
    raise ValueError("❌ Erro: SUPABASE_URL ou SUPABASE_KEY não encontradas no arquivo .env!")

supabase: Client = create_client(url, key)

# Variáveis globais para armazenar os dados da imagem selecionada
imagem_dados = None
imagem_nome = None

# --- FUNÇÕES ---
def upload_file():
    global imagem_dados, imagem_nome
    
    # Abre o gerenciador de arquivos filtrando por imagens
    file_path = filedialog.askopenfilename(
        filetypes=[("Imagens", "*.jpg *.jpeg *.png *.webp")]
    )
    
    if file_path:
        try:
            # Lê o arquivo em modo binário
            with open(file_path, 'rb') as f:
                imagem_dados = f.read()
            
            imagem_nome = os.path.basename(file_path)
            print(f"✅ Arquivo pronto para upload: {imagem_nome}")
            
            # Altera o texto do botão para mostrar que a imagem foi selecionada
            btn_upload.config(text=f"Selecionado: {imagem_nome[:15]}...", fg="green")
            
        except Exception as e:
            print(f"❌ Erro ao ler arquivo: {e}")
            messagebox.showerror("Erro", f"Não foi possível processar a imagem: {e}")

def atualizar_proximo_id():
    try:
        resposta = supabase.table("Informatica").select("ID").order("ID", desc=True).limit(1).execute()

        entry_ID.config(state="normal")
        entry_ID.delete("1.0", "end")
        
        if resposta.data:
            ultimo_id = resposta.data[0]["ID"]
            proximo_id = int(ultimo_id) + 1
            entry_ID.insert("1.0", str(proximo_id))
        else:
            entry_ID.insert("1.0", "1")
            
        entry_ID.config(state="disabled")
    except Exception as e:
        print(f"Erro ao buscar o próximo ID: {e}")
def formatar_moeda(event):
    # Obtém o texto atual do campo de preço
    texto = entry_Preco.get().strip()
    
    # Remove tudo o que não for número
    apenas_numeros = "".join(c for c in texto if c.isdigit())
    
    # Se deletar tudo, mantém o padrão R$ 0,00
    if not apenas_numeros:
        apenas_numeros = "0"
        
    # Converte para inteiro e divide por 100 para simular os centavos
    valor_float = int(apenas_numeros) / 100
    
    # Formata o número no padrão brasileiro
    texto_formatado = f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    # Atualiza o campo de texto sem perder a posição do cursor
    entry_Preco.delete(0, tk.END)
    entry_Preco.insert(0, texto_formatado)

        
def salvar_dados():
    global imagem_dados, imagem_nome

    def obter_texto(campo):
        try:
            return campo.get("1.0", "end-1c").strip()
        except TypeError:
            return campo.get().strip()

    nome_campo = obter_texto(entry_Nome)
    categoria_campo = obter_texto(entry_Categoria)
    marca_campo = obter_texto(entry_Marca)
    modelo_campo = obter_texto(entry_Modelo)
    fornecedor_campo = obter_texto(entry_Fornecedor)
    qtde_campo = obter_texto(entry_Qtde_Estoque)
    preco_campo = obter_texto(entry_Preco)
    descricao_campo = obter_texto(entry_Descricao)

    if not nome_campo:
        print("❌ Erro: O campo Nome é obrigatório!")
        messagebox.showerror("Erro", "O campo Nome é obrigatório!")
        return

    # --- NOVA TRAVA: IMPEDE QUE FIQUE VAZIO (EMPTY) OU NULL NO BANCO ---
    if not imagem_dados or not imagem_nome:
        print("❌ Erro: A imagem do produto é obrigatória!")
        messagebox.showerror("Erro", "Você precisa selecionar uma imagem para o produto!")
        return

    try:
        url_imagem_publica = None
        
        # --- PROCESSA O UPLOAD DA IMAGEM PARA O STORAGE ---
        if imagem_dados and imagem_nome:
            bucket_name = "Imagens"  # ⚠️ Substitua pelo nome do seu Bucket no Supabase
            
            # Descobre o content-type correto de forma automática (ex: image/jpeg)
            content_type, _ = mimetypes.guess_type(imagem_nome)
            if not content_type:
                content_type = "image/jpeg"
             
            
            # ✨ Cria um nome totalmente único misturando letras/números ao nome da foto
            nome_unico_arquivo = f"{uuid.uuid4()}-{imagem_nome}"
            
            print(f"Enviando imagem única '{nome_unico_arquivo}' para o bucket '{bucket_name}'...")
            
            # Faz o upload usando o nome único gerado
            supabase.storage.from_(bucket_name).upload(
                file=imagem_dados,
                path=nome_unico_arquivo,  # 👈 Mudou aqui
                file_options={"content-type": content_type}
            )
            
            # Gera a URL pública baseada no nome único gerado
            url_imagem_publica = supabase.storage.from_(bucket_name).get_public_url(nome_unico_arquivo) # 👈 Mudou aqui

            print(f"✅ Upload concluído. URL: {url_imagem_publica}")

        # Nova validação de segurança caso o Supabase falhe em retornar a URL
        if not url_imagem_publica or url_imagem_publica.strip() == "":
            raise ValueError("A URL pública gerada para a foto está vazia.")

             # --- TRATAMENTO E FORMATAÇÃO DO CAMPO PREÇO ---
                # --- TRATAMENTO E FORMATAÇÃO DO CAMPO PREÇO ---
        preco_final = 0.0
        if preco_campo:
            # Remove R$, espaços e pontos de milhar, ajustando a vírgula decimal
            preco_limpo = preco_campo.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".").strip()
            try:
                preco_final = float(preco_limpo) if preco_limpo else 0.0
            except ValueError:
                preco_final = 0.0

        # Dicionário configurado para a tabela do Banco de Dados
        dados = {
            "Nome": nome_campo,
            "Categoria": categoria_campo,
            "Marca": marca_campo,
            "Modelo": modelo_campo,
            "Fornecedor": fornecedor_campo,
            "Quantidade de Estoque": int(qtde_campo) if qtde_campo else 0,
            "Preço": preco_final,  
            "Descrição": descricao_campo,
            "Foto de Material": url_imagem_publica 
        }

        
        resposta = supabase.table("Informatica").insert(dados).execute()
        
        print("✅ Dados registrados com sucesso no Supabase!")
        print(f"Registro criado: {resposta.data}")
        messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!")
        
        # Limpa os campos de texto
        for campo in [entry_Nome, entry_Categoria, entry_Marca, entry_Modelo, entry_Fornecedor, entry_Qtde_Estoque, entry_Preco, entry_Descricao]:
            try:
                campo.delete("1.0", "end")
            except Exception:
                campo.delete(0, "end")
                
        # Reseta o botão de upload e as variáveis da imagem
        imagem_dados = None
        imagem_nome = None
        btn_upload.config(text="Selecionar Arquivo", fg="black")
        
        # Atualiza a interface com o próximo número de ID livre
        atualizar_proximo_id()
        
    except ValueError as error_conversao:
        print(f"❌ Erro de digitação: Verifique se Estoque e Preço são números válidos. Detalhe: {error_conversao}")
        messagebox.showerror("Erro", "Verifique se Estoque e Preço possuem caracteres válidos.")
    except Exception as e:
        print(f"❌ Ocorreu um erro ao salvar no Supabase: {e}")
        messagebox.showerror("Erro", f"Erro ao salvar no banco de dados/storage: {e}")


# --- INTERFACE GRÁFICA (TKINTER) ---
janela = tk.Tk()
janela.title("Sistema de Cadastro de Produtos de Informática - Maycon Depolli dos Santos")
janela.geometry("1800x1800")

label_ID = tk.Label(text="Código do Produto:", font=18, fg="blue")
label_ID.grid(row=1, column=1, padx=40, pady=15, sticky='ew', columnspan=1)
entry_ID = tk.Text(height=1, width=5, state="disabled")
entry_ID.grid(row=2, column=1, padx=40, pady=15, sticky='ew', columnspan=1)

label_Nome = tk.Label(text="Nome:", font=18, fg="blue")
label_Nome.grid(row=1, column=2, padx=40, pady=15, sticky='ew', columnspan=1)
entry_Nome = tk.Text(height=1, width=60)
entry_Nome.grid(row=2, column=2, padx=40, pady=15, sticky='ew', columnspan=1)

label_Categoria = tk.Label(text="Categoria/SubCategoria:", font=18, fg="blue")
label_Categoria.grid(row=1, column=3, padx=40, pady=15, sticky='ew', columnspan=1)
entry_Categoria = tk.Text(height=1, width=60)
entry_Categoria.grid(row=2, column=3, padx=40, pady=15, sticky='ew', columnspan=1)

label_Marca = tk.Label(text="Marca:", font=18, fg="blue")
label_Marca.grid(row=3, column=1, padx=40, pady=15, sticky='ew', columnspan=1)
entry_Marca = tk.Text(height=1, width=10)
entry_Marca.grid(row=4, column=1, padx=40, pady=15, sticky='ew', columnspan=1)

label_Modelo = tk.Label(text="Modelo:", font=18, fg="blue")
label_Modelo.grid(row=3, column=2, padx=10, pady=15, sticky='ew', columnspan=1)
entry_Modelo = tk.Text(height=1, width=10)
entry_Modelo.grid(row=4, column=2, padx=100, pady=15, sticky='ew', columnspan=1)

label_Fornecedor = tk.Label(text="Fornecedor:", font=18, fg="blue")
label_Fornecedor.grid(row=3, column=3, padx=10, pady=15, sticky='ew', columnspan=1)
entry_Fornecedor = tk.Text(height=1, width=10)
entry_Fornecedor.grid(row=4, column=3, padx=100, pady=15, sticky='ew', columnspan=1)

label_Qtde_Estoque = tk.Label(text="Quantidade de Estoque:", font=18, fg="blue")
label_Qtde_Estoque.grid(row=5, column=1, padx=10, pady=15, sticky='ew', columnspan=1)
entry_Qtde_Estoque = tk.Text(height=1, width=10)
entry_Qtde_Estoque.grid(row=6, column=1, padx=60, pady=15, sticky='ew', columnspan=1)

# Substitua a criação do seu entry_Preco por estas linhas:
label_Preco = tk.Label(text="Preço:", font=18, fg="blue")
label_Preco.grid(row=5, column=2, padx=10, pady=15, sticky='ew')

entry_Preco = tk.Entry(font=18) # Usar Entry facilita o controle de máscaras
entry_Preco.insert(0, "R$ 0,00") # Define o valor inicial padrão
entry_Preco.grid(row=6, column=2, padx=200, pady=15, sticky='ew')

# Vincula a função para rodar toda vez que uma tecla for solta
entry_Preco.bind("<KeyRelease>", formatar_moeda)


label_ft_Material = tk.Label(text="Foto de Material:", font=18, fg="blue")
label_ft_Material.grid(row=5, column=3, padx=10, pady=15, sticky='ew', columnspan=1)
btn_upload = tk.Button(janela, text="Selecionar Arquivo", command=upload_file)
btn_upload.grid(row=6, column=3, padx=200, pady=10, sticky='ew', columnspan=1)

label_Descricao = tk.Label(text="Descrição:", font=18, fg="blue")
label_Descricao.grid(row=8, column=1, padx=10, pady=15, sticky='ew', columnspan=3)
entry_Descricao = tk.Text(height=10, width=10)
entry_Descricao.grid(row=9, column=1, padx=10, pady=15, sticky='ew', columnspan=3)

try:
    bg_image = PhotoImage(file=r"Projeto\Outro Projeto\Maycon.jpg") 
except Exception as e:
    print(f"Erro ao carregar a imagem: {e}")
    bg_image = None

if bg_image:
    label_bg = tk.Label(janela, image=bg_image)
    label_bg.place(x=590, y=660)
    label_bg.image = bg_image 

label_Quem_fez = tk.Label(text="Desenvolvido por: Maycon Depolli dos Santos", font=18, fg="blue")
label_Quem_fez.grid(row=10, column=1, padx=50, pady=220, sticky='ew', columnspan=5)

botao_confirmar = tk.Button(janela, text="Confirmar", width=20, height=2, command=salvar_dados)
botao_confirmar.place(x=590, y=570)

atualizar_proximo_id()

janela.mainloop()
