import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --- SUA FUNÇÃO AQUI ---
def processar_planilha(caminho_excel):
    """
    Coloque aqui a lógica do seu código que processa a planilha.
    """
    print(f"Processando arquivo: {caminho_excel}")
    # Exemplo: df = pd.read_excel(caminho_excel)
    # ... seu código ...
    
    # Exemplo de retorno no final do processo
    return True

# --- INTERFACE GRÁFICA SIMPLES ---
def selecionar_e_rodar():
    # Esconde a janela principal do tkinter para mostrar só o leitor de arquivos
    root = tk.Tk()
    root.withdraw()

    # Abre a caixa de diálogo para escolher o arquivo .xlsx ou .xls
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha do Excel",
        filetypes=[("Arquivos do Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )

    # Se o usuário escolheu um arquivo (não cancelou)
    if caminho_arquivo:
        try:
            # Chama a sua função passando o caminho selecionado
            processar_planilha(caminho_arquivo)
            messagebox.showinfo("Sucesso!", "Planilha processada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar:\n{e}")
    else:
        print("Nenhum arquivo foi selecionado.")

if __name__ == "__main__":
    selecionar_e_rodar()
