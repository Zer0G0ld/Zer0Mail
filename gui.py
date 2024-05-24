import tkinter as tk
from tkinter import messagebox, simpledialog
from tkinter import ttk
from file_utils import *
from email_utils import send_email
from datetime import datetime
import os

listbox_toners = None  # Definindo a variável listbox_toners como global

def abrir_relatorio():
    print("DEBUG: Tentando abrir o relatório")
    if os.path.exists('relatorio_toner.xlsx'):
        os.system('xdg-open relatorio_toner.xlsx')
    else:
        messagebox.showerror("Erro", "Relatório não encontrado!")

def enviar_email():
    print("DEBUG: Enviar e-mail iniciado")
    to_email = simpledialog.askstring("Enviar E-mail", "Digite o e-mail do destinatário:")
    print(f"DEBUG: E-mail destinatário: {to_email}")
    if not to_email:
        messagebox.showerror("Erro", "Por favor, insira um e-mail!")
        return

    from_email = 'matheus.torres@reals.com.br'
    subject = 'Relatório de Toner'
    body = 'Por favor, encontre em anexo o relatório de toner.'
    data_atual = datetime.now().strftime('%Y-%m-%d')
    relatorio_filename = f'relatorio_toner_{data_atual}.xlsx'

    csv_to_excel('relatorio_toner.csv', relatorio_filename)

    try:
        send_email(from_email, to_email, subject, body, relatorio_filename)
        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
    except Exception as e:
        print(f"DEBUG: Erro ao enviar e-mail: {e}")
        messagebox.showerror("Erro", f"Erro ao enviar e-mail: {e}")

def adicionar_impressora_toner():
    print("DEBUG: Adicionar impressora ou toner iniciado")
    add_impressora_toner()
    messagebox.showinfo("Sucesso", "Impressora ou Toner adicionado com sucesso!")

def atualizar_tabela():
    print("DEBUG: Atualizar tabela iniciado")
    toner = simpledialog.askstring("Atualizar Tabela", "Digite o nome do toner a ser atualizado:")
    if toner:
        update_table(toner)
        messagebox.showinfo("Sucesso", "Tabela atualizada com sucesso!")
        visualizar_tabela()
    else:
        messagebox.showerror("Erro", "Nenhum toner informado.")

def quantificar():
    print("DEBUG: Quantificar toners iniciado")
    try:
        df = quantificar_toners()
        print(f"DEBUG: DataFrame retornado: {df}")
        if df is not None:
            messagebox.showinfo("Quantificação de Toners", df.to_string())
        else:
            messagebox.showerror("Erro", "Nenhum dado de quantificação encontrado.")
    except FileNotFoundError:
        print("DEBUG: Arquivo não encontrado")
        messagebox.showerror("Erro", "Nenhum relatório encontrado.")

def relacionar():
    print("DEBUG: Relacionar toners iniciado")
    impressora = simpledialog.askstring("Relacionar Toners", "Digite o nome da impressora:")
    if impressora:
        try:
            df = relacionar_toners_impressoras(impressora)
            print(f"DEBUG: DataFrame retornado: {df}")
            if df is not None:
                messagebox.showinfo("Relacionamento de Toners", df.to_string())
            else:
                messagebox.showerror("Erro", "Nenhum dado de relacionamento encontrado.")
        except FileNotFoundError:
            print("DEBUG: Arquivo não encontrado")
            messagebox.showerror("Erro", "Nenhum relatório encontrado.")
    else:
        messagebox.showerror("Erro", "Nenhuma impressora informada.")

def visualizar():
    print("DEBUG: Visualizar relatório iniciado")
    try:
        df = visualizar_relatorio()
        print(f"DEBUG: DataFrame retornado: {df}")
        if df is not None:
            messagebox.showinfo("Relatório de Toners", df.to_string())
        else:
            messagebox.showerror("Erro", "Nenhum relatório encontrado.")
    except FileNotFoundError:
        print("DEBUG: Arquivo não encontrado")
        messagebox.showerror("Erro", "Nenhum relatório encontrado.")

def visualizar_tabela():
    print("DEBUG: Visualizar tabela iniciado")
    df = visualizar_relatorio()
    print(f"DEBUG: DataFrame retornado: {df}")
    if df is not None:
        listbox_toners.delete(0, tk.END)
        for index, row in df.iterrows():
            listbox_toners.insert(tk.END, f"{row['Toner']}: {row['Quantidade']} unidades")
    else:
        messagebox.showerror("Erro", "Nenhum relatório encontrado.")

def start_gui():
    global listbox_toners  # Referenciar a variável global listbox_toners

    root = tk.Tk()
    root.title("StockPrint")

    # Função para maximizar a janela em tela cheia
    def toggle_fullscreen(event=None):
        root.attributes("-fullscreen", not root.attributes("-fullscreen"))

    root.bind("<F11>", toggle_fullscreen)  # Atalho de teclado para alternar tela cheia
    root.bind("<Escape>", toggle_fullscreen)  # Atalho de teclado para sair da tela cheia

    # Frame principal
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Título
    label_title = ttk.Label(main_frame, text="StockPrint - Controle de Toners", font=("Arial", 16, "bold"))
    label_title.pack(pady=(0, 20))

    # Frame dos botões
    frame_buttons = ttk.Frame(main_frame)
    frame_buttons.pack(side=tk.LEFT, fill='y', padx=10)

    # Botões
    button_texts = ["Adicionar Impressora/Toner", "Atualizar Tabela", "Quantificar Toners",
                    "Relacionar Toners a Impressoras", "Visualizar Relatório", "Abrir Relatório"]
    button_commands = [adicionar_impressora_toner, atualizar_tabela, quantificar, relacionar, visualizar, abrir_relatorio]

    for text, command in zip(button_texts, button_commands):
        btn = ttk.Button(frame_buttons, text=text, command=command, width=30)
        btn.pack(fill='x', pady=5)

    # Frame da tabela de toners
    frame_table = ttk.Frame(main_frame)
    frame_table.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)

    label_table = ttk.Label(frame_table, text="Tabela de Toners", font=("Arial", 12, "bold"))
    label_table.pack()

    scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL)
    listbox_toners = tk.Listbox(frame_table, yscrollcommand=scrollbar.set, font=("Arial", 10), width=40)
    scrollbar.config(command=listbox_toners.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox_toners.pack(fill=tk.BOTH, expand=True)

    visualizar_tabela()

    root.mainloop()

start_gui()
