"""
Conversor de Arquivo Word (.docx) para PDF com Interface Gráfica
----------------------------------------------------------------
Este programa permite ao usuário:
- Selecionar um arquivo Word (.docx)
- Escolher onde salvar o arquivo convertido
- Converter o documento para PDF usando a biblioteca docx2pdf

A interface gráfica é feita com tkinter.
Ideal para facilitar a conversão de documentos sem necessidade de abrir o Word.
"""

# Importa o módulo tkinter para a interface gráfica
import tkinter as tk
from tkinter import filedialog, messagebox

# Importa a função de conversão da biblioteca docx2pdf
from docx2pdf import convert

# Importa o módulo os para manipular nomes de arquivos e caminhos
import os

# Função que permite o usuário escolher o arquivo .docx a ser convertido
def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo Word",  # Título da janela
        filetypes=[("Documentos Word", "*.docx")]  # Filtro de arquivos: apenas .docx
    )
    # Se o usuário selecionou um arquivo, atualiza o campo de entrada
    if caminho:
        entrada_var.set(caminho)

# Função que permite escolher onde salvar o arquivo convertido
def escolher_local_salvar():
    caminho_entrada = entrada_var.get()
    # Verifica se um arquivo .docx foi selecionado
    if not caminho_entrada or not caminho_entrada.endswith('.docx'):
        messagebox.showerror("Erro", "Primeiro selecione um arquivo .docx válido.")
        return

    # Janela para salvar o arquivo, com sugestão de nome igual ao original
    caminho_pdf = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=os.path.splitext(os.path.basename(caminho_entrada))[0] + ".pdf"
    )
    # Se o usuário escolheu onde salvar, atualiza o campo de saída
    if caminho_pdf:
        saida_var.set(caminho_pdf)

# Função principal que realiza a conversão
def converter_para_pdf():
    caminho_docx = entrada_var.get()
    caminho_pdf = saida_var.get()

    # Verifica se o caminho de entrada é válido
    if not caminho_docx or not caminho_docx.endswith('.docx'):
        messagebox.showerror("Erro", "Selecione um arquivo .docx válido.")
        return

    # Verifica se o usuário escolheu onde salvar o PDF
    if not caminho_pdf:
        messagebox.showerror("Erro", "Escolha onde salvar o arquivo PDF.")
        return

    try:
        # Realiza a conversão usando a função convert da docx2pdf
        convert(caminho_docx, caminho_pdf)
        # Exibe mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Arquivo convertido e salvo como:\n{caminho_pdf}")
    except Exception as e:
        # Em caso de erro na conversão
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

# ------------------ INTERFACE GRÁFICA ------------------

# Cria a janela principal
janela = tk.Tk()
janela.title("Conversor Word para PDF")
janela.geometry("450x240")  # Define o tamanho da janela

# Variáveis de controle dos campos de entrada e saída
entrada_var = tk.StringVar()
saida_var = tk.StringVar()

# Texto e campo de entrada do caminho do arquivo Word
tk.Label(janela, text="Arquivo Word (.docx):").pack(pady=5)
tk.Entry(janela, textvariable=entrada_var, width=55).pack()
tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo).pack(pady=5)

# Texto e campo de entrada do local onde o PDF será salvo
tk.Label(janela, text="Salvar PDF como:").pack(pady=5)
tk.Entry(janela, textvariable=saida_var, width=55).pack()
tk.Button(janela, text="Escolher local para salvar", command=escolher_local_salvar).pack(pady=5)

# Botão final para realizar a conversão
tk.Button(
    janela,
    text="Converter para PDF",
    command=converter_para_pdf,
    bg="green", fg="white",
    width=25
).pack(pady=15)

# Inicia o loop da interface gráfica
janela.mainloop()
