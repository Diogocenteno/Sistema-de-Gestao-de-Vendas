#%%


import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import os
from ttkbootstrap import Style, themes
from datetime import datetime
import subprocess
import ctypes # Importa para manipular atributos de arquivo no Windows

# --- Classe para Gerenciamento do Banco de Dados ---
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._initialize_db()

    def _initialize_db(self):
        """Inicializa a conexão com o banco de dados e cria a tabela se não existir."""
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS vendas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_cliente TEXT,
                nome_produto TEXT,
                preco REAL,
                tipo_pagamento TEXT,
                preco_final REAL,
                nome_vendedor TEXT,
                data_hora TEXT
            )
            """)
            self.conn.commit()
            # Ocultar o arquivo do banco de dados (específico para Windows)
            # Este recurso tenta ocultar o arquivo .db para o usuário.
            # Em outros sistemas operacionais, esta função será ignorada.
            if os.name == 'nt':
                try:
                    if os.path.exists(self.db_path):
                        # Define o atributo de oculto (0x02)
                        ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x02)
                except Exception as e:
                    print(f"Não foi possível ocultar o arquivo do banco de dados: {e}")
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao inicializar o banco de dados: {e}\nVerifique as permissões de arquivo ou se o banco de dados está corrompido.")
            exit()

    def insert_sale(self, data):
        """Insere uma nova venda no banco de dados."""
        try:
            self.cursor.execute("""
            INSERT INTO vendas (nome_cliente, nome_produto, preco, tipo_pagamento, preco_final, nome_vendedor, data_hora)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                data["nome_cliente"], data["nome_produto"], data["preco"], data["tipo_pagamento"],
                data["preco_final"], data["nome_vendedor"], data["data_hora"]
            ))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao inserir venda: {e}\nPor favor, tente novamente. Se o problema persistir, reinicie a aplicação.")
            self.conn.rollback()

    def update_sale(self, sale_id, data):
        """Atualiza uma venda existente no banco de dados."""
        try:
            self.cursor.execute("""
            UPDATE vendas
            SET nome_cliente=?, nome_produto=?, preco=?, tipo_pagamento=?, preco_final=?, nome_vendedor=?, data_hora=?
            WHERE id=?
            """, (
                data["nome_cliente"], data["nome_produto"], data["preco"], data["tipo_pagamento"],
                data["preco_final"], data["nome_vendedor"], data["data_hora"], sale_id
            ))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao atualizar venda: {e}\nNão foi possível salvar as alterações. Tente novamente.")
            self.conn.rollback()

    def delete_sale(self, sale_id):
        """Exclui uma venda do banco de dados pelo ID."""
        try:
            self.cursor.execute("DELETE FROM vendas WHERE id=?", (sale_id,))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao excluir venda: {e}\nNão foi possível remover o registro. Tente novamente.")
            self.conn.rollback()

    def fetch_all_sales(self, search_term=""):
        """
        Busca todas as vendas do banco de dados, opcionalmente filtrando por um termo de busca.
        O termo de busca é aplicado aos campos nome_cliente, nome_produto e nome_vendedor.
        """
        try:
            query = "SELECT id, data_hora, nome_cliente, nome_produto, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas"
            params = []
            if search_term:
                search_pattern = f"%{search_term}%"
                query += " WHERE nome_cliente LIKE ? OR nome_produto LIKE ? OR nome_vendedor LIKE ?"
                params = [search_pattern, search_pattern, search_pattern]
            query += " ORDER BY id DESC"
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao buscar vendas: {e}\nNão foi possível carregar os dados da tabela. Verifique a conexão com o banco de dados.")
            return []

    def fetch_sale_by_id(self, sale_id):
        """Busca uma venda específica pelo ID."""
        try:
            self.cursor.execute("SELECT id, nome_cliente, nome_produto, preco, tipo_pagamento, nome_vendedor FROM vendas WHERE id=?", (sale_id,))
            return self.cursor.fetchone()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao buscar venda por ID: {e}\nNão foi possível recuperar os detalhes da venda para edição.")
            return None

    def close_connection(self):
        """Fecha a conexão com o banco de dados e reexibe o arquivo se estiver no Windows."""
        if self.conn:
            self.conn.close()
        if os.name == 'nt' and os.path.exists(self.db_path):
            try:
                # Remove o atributo de oculto (0x80)
                ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x80)
            except Exception as e:
                print(f"Não foi possível reexibir o arquivo do banco de dados: {e}")

# --- Nova Classe para o Caderno Virtual ---
class CadernoVirtual:
    # Delimitador para salvar os dados no arquivo de texto
    DATA_DELIMITER = "|" 
    TOTAL_ROW_ID = "caderno_total_row"

    def __init__(self, parent_root, theme_style, file_path, anotacoes_file_path):
        self.file_path = file_path
        self.anotacoes_file_path = anotacoes_file_path # Novo caminho para o arquivo de anotações
        self.caderno_window = tk.Toplevel(parent_root)
        self.caderno_window.title("Caderno de Encomendas")
        self.caderno_window.geometry("700x600") # Aumenta a janela para acomodar as colunas
        self.caderno_window.transient(parent_root) # Faz a janela do caderno ser filha da principal
        self.caderno_window.grab_set() # Impede interação com a janela pai enquanto o caderno está aberto
        self.caderno_window.protocol("WM_DELETE_WINDOW", self._on_closing) # Garante que o conteúdo seja salvo ao fechar

        # Aplicar o tema atual da aplicação principal
        self.style = theme_style
        self.style.configure('Caderno.TFrame', background=self.style.lookup('TFrame', 'background'))
        self.style.configure('Caderno.TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Caderno.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Caderno.Treeview.Heading', font=('Arial', 10, 'bold'))
        self.style.configure('Caderno.Treeview', rowheight=25)
        self.style.map('Caderno.Treeview', background=[('selected', self.style.colors.primary)])

        frame_caderno = ttk.Frame(self.caderno_window, padding=10, style='Caderno.TFrame')
        frame_caderno.pack(fill=tk.BOTH, expand=True)

        # Título
        ttk.Label(frame_caderno, text="Suas Encomendas", font=('Arial', 16, 'bold'), style='Caderno.TLabel').pack(pady=(0, 10))

        # Frame para os campos de entrada de nova encomenda
        frame_input = ttk.LabelFrame(frame_caderno, text="Adicionar/Editar Encomenda", padding=10)
        frame_input.pack(fill=tk.X, pady=(0, 10))

        # Campos de entrada
        ttk.Label(frame_input, text="Nome:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.nome_entry = ttk.Entry(frame_input, width=30)
        self.nome_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(frame_input, text="Produto:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.produto_entry = ttk.Entry(frame_input, width=30)
        self.produto_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(frame_input, text="Valor (R$):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.valor_entry = ttk.Entry(frame_input, width=30)
        self.valor_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        frame_input.grid_columnconfigure(1, weight=1)

        # Botões de ação para encomenda
        frame_botoes_input = ttk.Frame(frame_input)
        frame_botoes_input.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")

        self.btn_add_encomenda = ttk.Button(frame_botoes_input, text="Adicionar Encomenda", command=self._add_encomenda, style='success.Caderno.TButton')
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

        self.btn_update_encomenda = ttk.Button(frame_botoes_input, text="Atualizar Encomenda", command=self._update_encomenda, style='primary.Caderno.TButton')
        self.btn_update_encomenda.pack_forget() # Esconde inicialmente

        self.btn_cancel_edit = ttk.Button(frame_botoes_input, text="Cancelar Edição", command=self._reset_input_fields, style='secondary.Caderno.TButton')
        self.btn_cancel_edit.pack_forget() # Esconde inicialmente

        # Treeview para exibir as encomendas
        cols = ("Nome", "Produto", "Valor")
        self.tree = ttk.Treeview(frame_caderno, columns=cols, show="headings", height=10, style='Caderno.Treeview')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Configuração das colunas
        self.tree.heading("Nome", text="Nome do Cliente")
        self.tree.heading("Produto", text="Produto Encomendado")
        self.tree.heading("Valor", text="Valor (R$)")

        self.tree.column("Nome", width=200, anchor="center")
        self.tree.column("Produto", width=250, anchor="center")
        self.tree.column("Valor", width=100, anchor="center") # Alinha à direita para valores

        # Scrollbars
        scrollbar_y = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(self.tree, orient="horizontal", command=self.tree.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        # Evento de seleção na Treeview para edição
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.selected_item_iid = None # Para controlar o item selecionado para edição

        # Botões de gestão do caderno e o novo botão "Anotações"
        frame_botoes_gestao = ttk.Frame(frame_caderno, style='Caderno.TFrame')
        frame_botoes_gestao.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(frame_botoes_gestao, text="Excluir Encomenda Selecionada", command=self._delete_encomenda, style='danger.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes_gestao, text="Limpar Caderno Completo", command=self._clear_content, style='warning.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        
        # NOVO BOTÃO PARA ABRIR A JANELA DE ANOTAÇÕES
        ttk.Button(frame_botoes_gestao, text="Abrir Anotações", command=self._abrir_anotacoes, style='info.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        
        ttk.Button(frame_botoes_gestao, text="Fechar Caderno", command=self._on_closing, style='secondary.Caderno.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

        self._load_content()
        self._calculate_total() # Calcula o total inicial
        self.caderno_window.mainloop() # Inicia o loop para a nova janela

    def _load_content(self):
        """Carrega o conteúdo do arquivo para a Treeview."""
        self.tree.delete(*self.tree.get_children()) # Limpa a Treeview antes de carregar
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        parts = line.strip().split(self.DATA_DELIMITER)
                        if len(parts) == 3: # Espera 3 partes: Nome, Produto, Valor
                            # Tenta formatar o valor para exibição
                            try:
                                valor = float(parts[2].replace(",", "."))
                                parts[2] = f"{valor:.2f}".replace(".", ",")
                            except ValueError:
                                parts[2] = parts[2] # Mantém como está se não for um número válido
                            self.tree.insert('', tk.END, values=parts)
            except Exception as e:
                messagebox.showerror("Erro ao Carregar Caderno", f"Não foi possível carregar o caderno de encomendas: {e}")
        else:
            # Cria o arquivo se ele não existir
            try:
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.write("") # Garante que o arquivo esteja vazio inicialmente
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo do caderno: {e}")
        self._calculate_total()

    def _save_content(self):
        """Salva o conteúdo da Treeview para o arquivo."""
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                for item_id in self.tree.get_children():
                    if item_id == self.TOTAL_ROW_ID: # Não salva a linha de total
                        continue
                    values = self.tree.item(item_id, 'values')
                    # Converte o valor de volta para o formato numérico para salvar
                    # (removendo a vírgula e substituindo por ponto)
                    valor_salvar = values[2].replace(",", ".") if values[2] else "0.00"
                    
                    # Garante que o valor seja um float antes de salvar
                    try:
                        float(valor_salvar)
                    except ValueError:
                        valor_salvar = "0.00" # Define um valor padrão se não for numérico

                    f.write(f"{values[0]}{self.DATA_DELIMITER}{values[1]}{self.DATA_DELIMITER}{valor_salvar}\n")
            # messagebox.showinfo("Salvo", "Encomendas salvas com sucesso!") # Pode ser muito intrusivo
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Caderno", f"Não foi possível salvar o caderno de encomendas: {e}")
        self._calculate_total()


    def _add_encomenda(self):
        """Adiciona uma nova encomenda à Treeview e salva."""
        nome = self.nome_entry.get().strip()
        produto = self.produto_entry.get().strip()
        valor_str = self.valor_entry.get().strip()

        if not nome or not produto or not valor_str:
            messagebox.showwarning("Atenção", "Todos os campos (Nome, Produto, Valor) são obrigatórios.")
            return

        try:
            valor = float(valor_str.replace(",", "."))
            if valor < 0:
                messagebox.showwarning("Atenção", "O valor não pode ser negativo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Valor inválido! Digite um número válido (ex: 10,50 ou 10.50).")
            return

        # Formata o valor para exibição na Treeview
        valor_exibicao = f"{valor:.2f}".replace(".", ",")
        
        self.tree.insert('', tk.END, values=(nome, produto, valor_exibicao))
        self._save_content()
        self._reset_input_fields()
        self._calculate_total()

    def _on_tree_select(self, event):
        """Carrega os dados da encomenda selecionada para os campos de entrada para edição."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.selected_item_iid = None
            self._reset_input_fields()
            return

        # Ignora a linha de total
        if selected_items[0] == self.TOTAL_ROW_ID:
            self.tree.selection_remove(selected_items[0]) # Desseleciona
            self.selected_item_iid = None
            self._reset_input_fields()
            return

        self.selected_item_iid = selected_items[0]
        values = self.tree.item(self.selected_item_iid, 'values')

        self.nome_entry.delete(0, tk.END)
        self.nome_entry.insert(0, values[0])
        self.produto_entry.delete(0, tk.END)
        self.produto_entry.insert(0, values[1])
        self.valor_entry.delete(0, tk.END)
        self.valor_entry.insert(0, values[2]) # Mantém a vírgula para edição

        self.btn_add_encomenda.pack_forget()
        self.btn_update_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        self.btn_cancel_edit.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

    def _update_encomenda(self):
        """Atualiza a encomenda selecionada na Treeview e salva."""
        if not self.selected_item_iid:
            messagebox.showwarning("Atenção", "Nenhuma encomenda selecionada para atualizar.")
            return

        nome = self.nome_entry.get().strip()
        produto = self.produto_entry.get().strip()
        valor_str = self.valor_entry.get().strip()

        if not nome or not produto or not valor_str:
            messagebox.showwarning("Atenção", "Todos os campos (Nome, Produto, Valor) são obrigatórios.")
            return

        try:
            valor = float(valor_str.replace(",", "."))
            if valor < 0:
                messagebox.showwarning("Atenção", "O valor não pode ser negativo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Valor inválido! Digite um número válido (ex: 10,50 ou 10.50).")
            return
        
        valor_exibicao = f"{valor:.2f}".replace(".", ",")

        self.tree.item(self.selected_item_iid, values=(nome, produto, valor_exibicao))
        self._save_content()
        self._reset_input_fields()
        messagebox.showinfo("Sucesso", "Encomenda atualizada com sucesso!")
        self._calculate_total()

    def _delete_encomenda(self):
        """Exclui a(s) encomenda(s) selecionada(s) da Treeview e salva."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Atenção", "Selecione uma ou mais encomendas para excluir.")
            return

        if self.TOTAL_ROW_ID in selected_items:
            messagebox.showwarning("Atenção", "A linha de total não pode ser excluída.")
            return

        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir {len(selected_items)} encomenda(s) selecionada(s)? Esta ação não pode ser desfeita."):
            try:
                for item_id in selected_items:
                    self.tree.delete(item_id)
                self._save_content()
                self._reset_input_fields()
                messagebox.showinfo("Sucesso", f"{len(selected_items)} encomenda(s) excluída(s) com sucesso!")
                self._calculate_total()
            except Exception as e:
                messagebox.showerror("Erro ao Excluir", f"Ocorreu um erro ao excluir encomenda(s): {e}")

    def _reset_input_fields(self):
        """Limpa os campos de entrada e redefine os botões para o estado de adição."""
        self.nome_entry.delete(0, tk.END)
        self.produto_entry.delete(0, tk.END)
        self.valor_entry.delete(0, tk.END)
        self.selected_item_iid = None
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        self.btn_update_encomenda.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.tree.selection_remove(self.tree.selection()) # Desseleciona qualquer item

    def _clear_content(self):
        """Limpa todo o conteúdo do caderno após confirmação."""
        if messagebox.askyesno("Confirmar Limpar", "Tem certeza que deseja limpar todo o conteúdo do caderno? Esta ação não pode ser desfeita."):
            self.tree.delete(*self.tree.get_children()) # Limpa a Treeview
            self._save_content() # Salva o arquivo vazio
            self._reset_input_fields()
            messagebox.showinfo("Caderno Limpo", "O caderno foi limpo.")
            self._calculate_total()

    def _calculate_total(self):
        """Calcula e exibe o total dos valores das encomendas na Treeview."""
        # Remove a linha de total existente, se houver
        if self.tree.exists(self.TOTAL_ROW_ID):
            self.tree.delete(self.TOTAL_ROW_ID)

        total_valor = 0.0
        for item_id in self.tree.get_children():
            # Pula a linha de total se ela ainda existir por algum motivo (segurança)
            if item_id == self.TOTAL_ROW_ID:
                continue
            
            values = self.tree.item(item_id, 'values')
            if len(values) == 3:
                try:
                    # Remove a vírgula para converter para float
                    valor_str = values[2].replace(",", ".")
                    total_valor += float(valor_str)
                except ValueError:
                    pass # Ignora valores não numéricos

        total_formatado = f"{total_valor:.2f}".replace(".", ",")
        self.tree.insert('', tk.END, iid=self.TOTAL_ROW_ID, values=(
            "", "TOTAL GERAL:", total_formatado
        ), tags=('total_row',))
        
        # Configura o estilo da linha de total
        self.style.configure('total_row.Caderno.Treeview', background='#e0e0e0', font=('Arial', 10, 'bold'))
        self.tree.tag_configure('total_row', background='#e0e0e0', font=('Arial', 10, 'bold'))


    def _abrir_anotacoes(self):
        """Abre a janela virtual para anotações."""
        # Passa também o caminho do arquivo de anotações para o CadernoVirtual
        AnotacoesVirtual(self.caderno_window, self.style, self.anotacoes_file_path)

    def _on_closing(self):
        """Salva o conteúdo antes de fechar a janela do caderno."""
        self._save_content() # Garante que as últimas alterações sejam salvas
        self.caderno_window.destroy()
        self.caderno_window.grab_release() # Libera o foco para a janela pai


# --- NOVA CLASSE PARA AS ANOTAÇÕES VIRTUAIS ---
class AnotacoesVirtual:
    def __init__(self, parent_root, theme_style, file_path):
        self.file_path = file_path
        self.anotacoes_window = tk.Toplevel(parent_root)
        self.anotacoes_window.title("Anotações")
        self.anotacoes_window.geometry("500x600")
        self.anotacoes_window.transient(parent_root)
        self.anotacoes_window.grab_set()
        self.anotacoes_window.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.style = theme_style
        self.style.configure('Anotacoes.TFrame', background=self.style.lookup('TFrame', 'background'))
        self.style.configure('Anotacoes.TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Anotacoes.TLabel', font=('Arial', 14, 'bold'))

        frame_anotacoes = ttk.Frame(self.anotacoes_window, padding=10, style='Anotacoes.TFrame')
        frame_anotacoes.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame_anotacoes, text="Suas Anotações", font=('Arial', 18, 'bold'), style='Anotacoes.TLabel').pack(pady=(0, 10))

        # Frame para conter o canvas e o texto
        frame_texto = ttk.Frame(frame_anotacoes)
        frame_texto.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Canvas para desenhar as linhas do caderno
        self.canvas = tk.Canvas(frame_texto, 
                                bg=self.style.lookup('TEntry', 'fieldbackground'),
                                highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar vertical
        # A scrollbar controlará a rolagem tanto do canvas quanto da área de texto.
        scrollbar_y = ttk.Scrollbar(frame_texto, orient="vertical", command=self._scroll_both) 
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Campo de texto sobre o canvas
        self.text_area = tk.Text(self.canvas, wrap=tk.WORD, font=('Courier New', 12),
                                 bg=self.style.lookup('TEntry', 'fieldbackground'),
                                 fg=self.style.lookup('TEntry', 'foreground'),
                                 insertbackground=self.style.lookup('TEntry', 'foreground'),
                                 relief=tk.FLAT, bd=2) 
        
        # Coloca o widget de texto no canvas na posição (0,0)
        self.text_window = self.canvas.create_window((0, 0), window=self.text_area, anchor="nw")
        
        # Conecta a rolagem interna da área de texto à barra de rolagem.
        # A barra de rolagem, por sua vez, tem seu comando configurado para rolar ambos os widgets.
        self.text_area.configure(yscrollcommand=scrollbar_y.set)

        # Configura eventos para redimensionamento e liberação de teclas
        # Quando o conteúdo ou tamanho da área de texto muda, ajusta a região de rolagem do canvas e redesenha as linhas.
        self.text_area.bind("<Configure>", self._ajustar_canvas)
        self.text_area.bind("<KeyRelease>", self._ajustar_canvas) 
        # Quando o próprio canvas é redimensionado, ajusta a largura e altura da janela da área de texto e redesenha as linhas.
        self.canvas.bind("<Configure>", self._ajustar_texto_e_linhas) 
        
        # Desenha as linhas iniciais
        self._desenhar_linhas()

        # Botões
        frame_botoes = ttk.Frame(frame_anotacoes, style='Anotacoes.TFrame')
        frame_botoes.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(frame_botoes, text="Salvar Anotações", command=self._save_content, style='info.Anotacoes.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Limpar Anotações", command=self._clear_content, style='danger.Anotacoes.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Fechar Anotações", command=self._on_closing, style='secondary.Anotacoes.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

        self._load_content()
        self.anotacoes_window.mainloop()

    def _scroll_both(self, *args):
        """Rola tanto o canvas quanto a área de texto."""
        self.canvas.yview(*args)
        self.text_area.yview(*args)
        # Redesenha as linhas para garantir que estejam consistentes com a posição de rolagem atual
        self._desenhar_linhas() 

    def _ajustar_canvas(self, event=None):
        """Ajusta a região de rolagem do canvas com base no conteúdo do text_area e redesenha as linhas."""
        # Obtém a altura total do conteúdo de texto
        try:
            # Obtém a caixa delimitadora do último caractere na área de texto
            # Isso nos dá a extensão do conteúdo de texto
            bbox_end = self.text_area.bbox(tk.END + '-1c') 
            if bbox_end:
                # bbox_end é (x, y, largura, altura)
                # A altura total do conteúdo é y + altura
                total_text_height = bbox_end[1] + bbox_end[3]
            else:
                # Se não houver texto, define a altura mínima para a altura do canvas
                total_text_height = self.canvas.winfo_height() 
        except tk.TclError: 
            # Lida com casos em que a área de texto pode estar vazia ou não pronta
            total_text_height = self.canvas.winfo_height()

        # Define a região de rolagem do canvas para cobrir todo o conteúdo de texto
        # A dimensão X da região de rolagem deve corresponder à largura do canvas
        self.canvas.config(scrollregion=(0, 0, self.canvas.winfo_width(), total_text_height))
        # Redesenha as linhas após ajustar o tamanho do canvas
        self._desenhar_linhas() 

    def _ajustar_texto_e_linhas(self, event=None):
        """Ajusta o widget de texto e redesenha as linhas quando o canvas é redimensionado."""
        # Ajusta a largura da janela da área de texto dentro do canvas para corresponder à largura do canvas
        self.canvas.itemconfig(self.text_window, width=event.width)
        # Ajusta a altura da janela da área de texto dentro do canvas para corresponder à altura do canvas
        # Isso faz com que a área de texto preencha a área visível do canvas
        self.canvas.itemconfig(self.text_window, height=event.height)
        # Redesenha as linhas após o redimensionamento do canvas
        self._desenhar_linhas()

    def _desenhar_linhas(self, event=None):
        """Desenha linhas no canvas para simular um caderno."""
        self.canvas.delete("linha")  # Remove linhas existentes
        
        # Obtém a altura de uma única linha de texto
        try:
            line_height_info = self.text_area.dlineinfo("1.0")
            if not line_height_info:
                return
            line_height = line_height_info[3]  # Altura em pixels
            if line_height <= 0:
                return
        except tk.TclError:
            return

        # Obtém a altura total rolável do canvas a partir de sua região de rolagem
        # Isso representa a altura total do conteúdo, incluindo as partes roladas para fora da vista.
        scroll_region_str = self.canvas.cget("scrollregion")
        if scroll_region_str:
            # scroll_region_str é uma string como "x1 y1 x2 y2"
            parts = scroll_region_str.split()
            if len(parts) == 4:
                total_drawable_height = float(parts[3]) # A coordenada y2 é a altura total
            else:
                total_drawable_height = self.canvas.winfo_height() # Fallback se a região de rolagem estiver malformada
        else:
            total_drawable_height = self.canvas.winfo_height() # Fallback se nenhuma região de rolagem estiver definida ainda

        # Desenha linhas do topo da região rolável até a sua parte inferior
        # Adiciona algumas linhas extras para garantir cobertura mesmo com pequenos desalinhamentos
        for i in range(int(total_drawable_height / line_height) + 5): 
            y = i * line_height
            self.canvas.create_line(0, y, self.canvas.winfo_width(), y, 
                                     fill="#cccccc", tags="linha")

    def _load_content(self):
        """Carrega o conteúdo do arquivo para o campo de texto."""
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.text_area.delete(1.0, tk.END)
                    self.text_area.insert(1.0, content)
            except Exception as e:
                messagebox.showerror("Erro ao Carregar Anotações", f"Não foi possível carregar as anotações: {e}")
        else:
            # Cria o arquivo se ele não existir
            try:
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.write("") # Garante que o arquivo esteja vazio inicialmente
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo de anotações: {e}")
        # Garante que as linhas sejam desenhadas após o carregamento do conteúdo
        self._ajustar_canvas() # Isso também chamará _desenhar_linhas

    def _save_content(self):
        """Salva o conteúdo do campo de texto para o arquivo."""
        try:
            content = self.text_area.get(1.0, tk.END).strip()
            with open(self.file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Salvo", "Anotações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Anotações", f"Não foi possível salvar as anotações: {e}")

    def _clear_content(self):
        """Limpa todo o conteúdo das anotações após confirmação."""
        if messagebox.askyesno("Confirmar Limpar", "Tem certeza que deseja limpar todo o conteúdo das anotações? Esta ação não pode ser desfeita."):
            self.text_area.delete(1.0, tk.END)
            self._save_content() # Salva o arquivo vazio
            messagebox.showinfo("Anotações Limpas", "As anotações foram limpas.")
            self._ajustar_canvas() # Atualiza as linhas após a limpeza

    def _on_closing(self):
        """Salva o conteúdo antes de fechar a janela das anotações."""
        self._save_content()
        self.anotacoes_window.destroy()
        self.anotacoes_window.grab_release()

# --- Classe Principal da Aplicação ---
class SalesApp:
    def __init__(self, root):
        self.root = root
        
        # Define o caminho base para o banco de dados e para os cadernos
        app_dir = os.path.dirname(os.path.abspath(__file__))
        internal_data_dir = os.path.join(app_dir, "_internal")
        
        self.db_path = os.path.join(internal_data_dir, "vendas_padaria.db")
        self.caderno_path = os.path.join(internal_data_dir, "encomendas_caderno.txt")
        self.anotacoes_path = os.path.join(internal_data_dir, "anotacoes.txt")
        self.theme_settings_path = os.path.join(internal_data_dir, "theme_setting.txt") # Novo caminho para o arquivo de tema

        # Verifica se o diretório _internal existe e é gravável, ou tenta criá-lo
        try:
            os.makedirs(internal_data_dir, exist_ok=True)
            if not os.access(internal_data_dir, os.W_OK):
                    raise OSError("Diretório '_internal' não é gravável.")
        except OSError:
            # Fallback para a pasta Documentos se não for possível usar _internal
            self.documents_path = os.path.join(os.path.expanduser("~"), "Documents")
            self.db_path = os.path.join(self.documents_path, "vendas_padaria.db")
            self.caderno_path = os.path.join(self.documents_path, "encomendas_caderno.txt")
            self.anotacoes_path = os.path.join(self.documents_path, "anotacoes.txt")
            self.theme_settings_path = os.path.join(self.documents_path, "theme_setting.txt") # ATUALIZA O CAMINHO DO TEMA
            os.makedirs(os.path.dirname(self.db_path), exist_ok=True) # Garante que a pasta Documentos exista
            messagebox.showwarning("Aviso de Caminho de Dados",
                                   f"Não foi possível criar ou acessar a pasta '_internal' no diretório do aplicativo.\n"
                                   f"O banco de dados, caderno e anotações serão salvos em: {self.documents_path}")

        # O caminho para a pasta de exportação de Excel permanece na pasta Documentos
        self.excel_export_folder = os.path.join(os.path.expanduser("~"), "Documents", "Vendas_Padaria_Exportadas")
        os.makedirs(self.excel_export_folder, exist_ok=True)

        # Inicializa o estilo com um tema padrão para que theme_names() esteja disponível
        self.style = Style(theme="cosmo") # Tema padrão inicial
        # Carrega o tema salvo ou define um padrão
        self.current_theme_name = self._load_theme_setting()
        # Aplica o tema carregado (ou o padrão)
        self.style.theme_use(self.current_theme_name)
        
        # --- Configuração global de estilo para botões arredondados ---
        # Aumenta o padding para dar uma aparência mais "cheia" e acentuar o arredondamento
        self.style.configure("TButton", padding=(10, 5)) 
        # --- Fim da configuração global de estilo ---

        self.root.title("Sistema de Cadastro de Vendas Padaria do Jeff e Pri")
        self.root.geometry("1050x800")

        self.db_manager = DatabaseManager(self.db_path)
        self.id_venda_em_edicao = None

        self._create_widgets()
        self.atualizar_tabela()
        self.limpar_campos_e_resetar_edicao()

        self.root.protocol("WM_DELETE_WINDOW", self.fechar_app)

    def _create_widgets(self):
        """Cria e organiza todos os widgets da interface gráfica."""
        # Título principal com fonte maior e negrito para um visual impactante (tamanho 36)
        title_label = ttk.Label(self.root, text="SISTEMA DE GESTÃO DE VENDAS", font=("Arial", 36, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 15))

        frame_theme_selector = ttk.Frame(self.root, padding=(10, 5))
        frame_theme_selector.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(5, 0))

        ttk.Label(frame_theme_selector, text="Escolher Tema:").pack(side=tk.RIGHT, padx=(0, 5))
        # Usa uma lista predefinida de temas para melhor controle e experiência do usuário
        self.available_themes = sorted(list(self.style.theme_names()))
        self.theme_combobox = ttk.Combobox(frame_theme_selector, values=self.available_themes, state="readonly", width=20)
        self.theme_combobox.set(self.current_theme_name)
        self.theme_combobox.pack(side=tk.RIGHT)
        self.theme_combobox.bind("<<ComboboxSelected>>", self.change_theme)

        frame_formulario = ttk.LabelFrame(self.root, text="Registo de Venda Padaria do Jeff e Pri", padding=(10,10))
        frame_formulario.grid(row=2, column=0, padx=10, pady=(5, 10), sticky="ew", columnspan=3)

        labels_info = [
            ("nome_cliente", "Nome do Cliente:"),
            ("nome_produto", "Nome do Produto:"),
            ("preco", "Preço (R$):"),
            ("tipo_pagamento", "Tipo de Pagamento:"),
            ("nome_vendedor", "Nome do Vendedor:"),
        ]
        self.campos = {}

        for i, (chave, texto_label) in enumerate(labels_info):
            label = ttk.Label(frame_formulario, text=texto_label)
            label.grid(row=i, column=0, sticky="w", padx=5, pady=5)
            
            if chave == "tipo_pagamento":
                combo_pagamento = ttk.Combobox(frame_formulario, values=["Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Pix"], state="readonly", width=28)
                combo_pagamento.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.campos[chave] = combo_pagamento
            else:
                entrada = ttk.Entry(frame_formulario, width=30)
                entrada.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.campos[chave] = entrada

        frame_formulario.grid_columnconfigure(1, weight=1)

        frame_botoes_formulario = ttk.Frame(frame_formulario)
        frame_botoes_formulario.grid(row=len(labels_info), column=0, columnspan=2, pady=10, sticky="ew")

        self.btn_salvar = ttk.Button(frame_botoes_formulario, text="Registar Venda", command=self.salvar_dados, style="success.TButton")
        self.btn_salvar.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.btn_cancelar_edicao = ttk.Button(frame_botoes_formulario, text="Cancelar Edição", command=self.limpar_campos_e_resetar_edicao, style="secondary.TButton")
        self.btn_cancelar_edicao.pack_forget()

        frame_botoes_gestao = ttk.LabelFrame(self.root, text="Gestão de Vendas e Exportação", padding=(10,10))
        frame_botoes_gestao.grid(row=3, column=0, padx=10, pady=5, sticky="ew", columnspan=3)

        ttk.Button(frame_botoes_gestao, text="Carregar para Editar", command=self.carregar_para_edicao, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Excluir Venda Selecionada", command=self.excluir_venda_selecionada, style="danger.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Exportar para Excel", command=self.exportar_excel, style="primary.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Abrir Planilha Excel", command=self.abrir_planilha_excel, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Limpar Tabela (Visual)", command=self.limpar_tabela_visual, style="warning.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Abrir Calculadora", command=self.abrir_calculadora, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        # O botão Encomendas agora chama a nova função para o caderno virtual
        ttk.Button(frame_botoes_gestao, text="Encomendas", command=self.abrir_caderno_encomendas, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        frame_tabela = ttk.LabelFrame(self.root, text="Vendas Registradas", padding=(10,10))
        frame_tabela.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        self.root.grid_rowconfigure(4, weight=1) 
        self.root.grid_columnconfigure(0, weight=1) 
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

        frame_tabela.grid_rowconfigure(1, weight=1)
        frame_tabela.grid_columnconfigure(0, weight=1)

        frame_pesquisa = ttk.Frame(frame_tabela, padding=(0, 5))
        frame_pesquisa.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        ttk.Label(frame_pesquisa, text="Pesquisar:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(frame_pesquisa, width=50)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.search_entry.bind("<KeyRelease>", self.on_search_key_release)

        ttk.Button(frame_pesquisa, text="Limpar Pesquisa", command=self.clear_search, style="secondary.TButton").pack(side=tk.LEFT, padx=(5, 0))

        colunas_visiveis = ("data_hora", "nome_cliente", "nome_produto", "preco", "tipo_pagamento", "preco_final", "nome_vendedor")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas_visiveis, show="headings", height=10, selectmode="extended") 
        self.tree.grid(row=1, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree.xview)
        scrollbar_x.grid(row=2, column=0, sticky="ew", columnspan=2)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.heading("data_hora", text="Data e Hora")
        self.tree.heading("nome_cliente", text="Cliente")
        self.tree.heading("nome_produto", text="Produto")
        self.tree.heading("preco", text="Preço (R$)")
        self.tree.heading("tipo_pagamento", text="Pagamento")
        self.tree.heading("preco_final", text="Preço Final (R$)")
        self.tree.heading("nome_vendedor", text="Vendedor")

        self.tree.column("data_hora", width=140, anchor="center", stretch=tk.NO)
        self.tree.column("nome_cliente", width=180, anchor="center", stretch=tk.YES) # Permite esticar
        self.tree.column("nome_produto", width=180, anchor="center", stretch=tk.YES) # Permite esticar
        self.tree.column("preco", width=90, anchor="center", stretch=tk.NO)
        self.tree.column("tipo_pagamento", width=120, anchor="center", stretch=tk.NO)
        self.tree.column("preco_final", width=100, anchor="center", stretch=tk.NO)
        self.tree.column("nome_vendedor", width=120, anchor="center", stretch=tk.YES) # Permite esticar

    def abrir_calculadora(self):
        """Abre a calculadora padrão do sistema operacional."""
        try:
            if os.name == 'nt': # Windows
                subprocess.Popen(['calc.exe'])
            elif os.uname().sysname == 'Darwin': # macOS
                subprocess.Popen(['open', '-a', 'Calculator'])
            else: # Linux/Outros Unix-like
                try:
                    subprocess.Popen(['gnome-calculator'])
                except FileNotFoundError:
                    try:
                        subprocess.Popen(['kcalc'])
                    except FileNotFoundError:
                        subprocess.Popen(['xcalc'])
        except FileNotFoundError:
            messagebox.showerror("Erro", "Calculadora não encontrada ou não configurada no sistema.\nPor favor, verifique se a calculadora está instalada e acessível no seu PATH.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir a calculadora: {e}\nVerifique se o aplicativo da calculadora está funcionando corretamente.")

    def abrir_caderno_encomendas(self):
        """Abre a janela do caderno virtual para anotações de encomendas."""
        # Passa também o caminho do arquivo de anotações para o CadernoVirtual
        CadernoVirtual(self.root, self.style, self.caderno_path, self.anotacoes_path)

    def abrir_planilha_excel(self):
        """Abre o arquivo Excel exportado, se existir."""
        file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Atenção", "O arquivo Excel ainda não foi exportado.\nPor favor, exporte-o primeiro clicando no botão 'Exportar para Excel'.")
            return
        
        try:
            if os.name == 'nt': # Windows
                os.startfile(file_path)
            elif os.uname().sysname == 'Darwin': # macOS
                subprocess.Popen(['open', file_path])
            else: # Linux/Outros Unix-like
                subprocess.Popen(['xdg-open', file_path])
        except FileNotFoundError:
            messagebox.showerror("Erro", "Nenhum aplicativo associado encontrado para abrir arquivos .xlsx.\nCertifique-se de ter o Microsoft Excel ou um programa compatível instalado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir a planilha: {e}\nVerifique se o arquivo Excel não está corrompido ou em uso por outro programa.")

    def salvar_dados(self):
        """
        Salva ou atualiza os dados de uma venda no banco de dados.
        Valida os campos obrigatórios e formata o preço.
        """
        try:
            dados = {
                "nome_cliente": self.campos["nome_cliente"].get().strip(),
                "nome_produto": self.campos["nome_produto"].get().strip(),
                "preco": self.campos["preco"].get().strip(),
                "tipo_pagamento": self.campos["tipo_pagamento"].get().strip(),
                "nome_vendedor": self.campos["nome_vendedor"].get().strip(),
            }

            # --- Validação de campos obrigatórios ---
            if not dados["nome_cliente"]:
                messagebox.showwarning("Atenção", "O campo 'Nome do Cliente' é obrigatório e não pode estar vazio.")
                self.campos["nome_cliente"].focus_set() # Foca no campo
                return
            if not dados["nome_produto"]:
                messagebox.showwarning("Atenção", "O campo 'Nome do Produto' é obrigatório e não pode estar vazio.")
                self.campos["nome_produto"].focus_set()
                return
            if not dados["preco"]:
                messagebox.showwarning("Atenção", "O campo 'Preço' é obrigatório.")
                self.campos["preco"].focus_set()
                return
            if not dados["tipo_pagamento"]:
                messagebox.showwarning("Atenção", "O campo 'Tipo de Pagamento' é obrigatório.\nPor favor, selecione uma opção na lista.")
                self.campos["tipo_pagamento"].focus_set()
                return
            if not dados["nome_vendedor"]:
                messagebox.showwarning("Atenção", "O campo 'Nome do Vendedor' é obrigatório e não pode estar vazio.")
                self.campos["nome_vendedor"].focus_set()
                return

            # --- Validação e conversão do preço ---
            try:
                # Remove espaços e tenta converter para float, aceitando tanto ',' quanto '.'
                preco_str = dados["preco"].replace(",", ".")
                preco = float(preco_str)
                if preco < 0:
                    messagebox.showwarning("Atenção", "O preço não pode ser um valor negativo.")
                    self.campos["preco"].focus_set()
                    return
            except ValueError:
                messagebox.showwarning("Atenção", "Preço inválido! Digite um número válido (ex: 10,50 ou 10.50).\nVerifique se não há letras ou múltiplos pontos/vírgulas.")
                self.campos["preco"].focus_set()
                return

            preco_final = preco # Manter como está, para futuras modificações de cálculo
            data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            dados["preco"] = preco
            dados["preco_final"] = preco_final
            dados["data_hora"] = data_hora_atual

            if self.id_venda_em_edicao is not None:
                self.db_manager.update_sale(self.id_venda_em_edicao, dados)
                messagebox.showinfo("Sucesso", "Venda atualizada com sucesso!")
            else:
                self.db_manager.insert_sale(dados)
                messagebox.showinfo("Sucesso", "Venda registrada com sucesso!")

            self.limpar_campos_e_resetar_edicao()
            self.atualizar_tabela()

        except Exception as e:
            # Garante que a mensagem de erro seja exibida corretamente mesmo com caracteres especiais
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Erro ao salvar dados: {error_message}\nOcorreu um problema inesperado ao tentar registrar/atualizar a venda. Verifique os dados inseridos.")


    def limpar_campos_e_resetar_edicao(self):
        """Limpa todos os campos do formulário e redefine o estado de edição."""
        self.id_venda_em_edicao = None
        for chave in self.campos:
            if isinstance(self.campos[chave], ttk.Combobox):
                self.campos[chave].set("")
            else:
                if self.campos[chave]:
                    self.campos[chave].delete(0, tk.END)

        self.btn_salvar.config(text="Registar Venda", style="success.TButton")
        self.btn_cancelar_edicao.pack_forget()

    def carregar_para_edicao(self):
        """
        Carrega os dados da venda selecionada na tabela para os campos de edição.
        Avisa se múltiplas seleções forem feitas e carrega apenas a primeira.
        """
        selecionado = self.tree.selection()
        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione uma venda na tabela para editar.")
            return
        
        if len(selecionado) > 1:
            messagebox.showwarning("Atenção", "Selecione apenas UMA venda para editar.\nApenas a primeira venda selecionada será carregada.")
            item_id_selecionado = selecionado[0]
        else:
            item_id_selecionado = selecionado[0]

        if item_id_selecionado == "total_row_id":
            messagebox.showwarning("Atenção", "A linha de total não pode ser editada.")
            return

        try:
            venda_id_db = int(item_id_selecionado) 
            venda = self.db_manager.fetch_sale_by_id(venda_id_db)

            if venda:
                self.limpar_campos_e_resetar_edicao()

                self.campos["nome_cliente"].insert(0, venda[1])
                self.campos["nome_produto"].insert(0, venda[2])
                self.campos["preco"].insert(0, str(venda[3]).replace(".", ","))
                self.campos["tipo_pagamento"].set(venda[4])
                self.campos["nome_vendedor"].insert(0, venda[5])
                
                self.id_venda_em_edicao = venda_id_db
                self.btn_salvar.config(text="Atualizar Venda", style="primary.TButton")
                
                self.btn_cancelar_edicao.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

            else:
                messagebox.showerror("Erro", "Venda não encontrada no banco de dados.\nO registro pode ter sido excluído ou o ID é inválido.")
                self.id_venda_em_edicao = None

        except ValueError:
            messagebox.showerror("Erro", "ID de venda inválido.\nSelecione uma linha de venda válida na tabela.")
            self.id_venda_em_edicao = None
        except Exception as e:
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Erro ao carregar dados para edição: {error_message}\nOcorreu um problema ao tentar carregar a venda selecionada. Tente novamente.")
            self.id_venda_em_edicao = None

    def excluir_venda_selecionada(self):
        """Exclui a(s) venda(s) selecionada(s) da tabela e do banco de dados."""
        selecionado = self.tree.selection()
        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione uma ou mais vendas na tabela para excluir.")
            return

        if "total_row_id" in selecionado:
            messagebox.showwarning("Atenção", "A linha de total não pode ser excluída.\nPor favor, desfaça a seleção da linha de total e tente novamente.")
            return

        if len(selecionado) > 1:
            confirmation_message = f"Tem certeza que deseja excluir as {len(selecionado)} vendas selecionadas? Esta ação não pode ser desfeita."
        else:
            confirmation_message = "Tem certeza que deseja excluir a venda selecionada? Esta ação não pode ser desfeita."
            
        if messagebox.askyesno("Confirmar Exclusão", confirmation_message):
            try:
                for item_id_selecionado in selecionado:
                    self.db_manager.delete_sale(int(item_id_selecionado))
                
                messagebox.showinfo("Sucesso", f"{len(selecionado)} venda(s) excluída(s) com sucesso!")
                self.atualizar_tabela()
                self.limpar_campos_e_resetar_edicao()
            except ValueError:
                messagebox.showerror("Erro", "ID de venda inválido.\nSelecione uma linha de venda válida para exclusão.")
            except Exception as e:
                error_message = str(e).replace('%', '%%')
                messagebox.showerror("Erro", f"Erro ao excluir venda(s): {error_message}\nOcorreu um problema ao tentar excluir a(s) venda(s) selecionada(s).")

    def exportar_excel(self):
        """Exporta todos os dados de vendas para um arquivo Excel."""
        if not messagebox.askyesno("Confirmar Exportação", "Deseja realmente exportar todas as vendas para um arquivo Excel?\nIsso pode levar alguns segundos se houver muitos dados."):
            return

        try:
            df = pd.read_sql_query("SELECT id, data_hora, nome_cliente, nome_produto, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas", self.db_manager.conn)
            if df.empty:
                messagebox.showinfo("Informação", "Não há vendas para exportar.")
                return

            total_vendas = df["preco_final"].sum()
            
            df.rename(columns={
                'id': 'ID da Venda',
                'data_hora': 'Data e Hora',
                'nome_cliente': 'Nome do Cliente',
                'nome_produto': 'Nome do Produto',
                'preco': 'Preço Unitário (R$)',
                'tipo_pagamento': 'Tipo de Pagamento',
                'nome_vendedor': 'Nome do Vendedor',
                'preco_final': 'Preço Final (R$)'
            }, inplace=True)

            total_row_df = pd.DataFrame([{
                "ID da Venda": "", 
                "Data e Hora": "",
                "Nome do Cliente": "",
                "Nome do Produto": "",
                "Preço Unitário (R$)": float('nan'), # Definido como NaN para compatibilidade com float
                "Tipo de Pagamento": "TOTAL GERAL:",
                "Preço Final (R$)": total_vendas,
                "Nome do Vendedor": ""
            }])
            
            df_final = pd.concat([df, total_row_df], ignore_index=True)

            os.makedirs(self.excel_export_folder, exist_ok=True)
            
            file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")

            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Vendas', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Vendas']

                money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

                worksheet.set_column(4, 4, None, money_format)
                worksheet.set_column(6, 6, None, money_format)

            messagebox.showinfo("Exportado", f"Arquivo Excel salvo com sucesso em:\n{file_path}\nTotal das vendas incluído e colunas formatadas!")
        except Exception as e:
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {error_message}\nVerifique se o arquivo não está aberto em outro programa ou se há espaço em disco suficiente.")

    def limpar_tabela_visual(self):
        """Limpa todas as linhas da Treeview (apenas visualmente, não do banco de dados)."""
        for row in self.tree.get_children():
            self.tree.delete(row)

    def change_theme(self, event=None):
        """Altera o tema da aplicação com base na seleção do combobox e salva a preferência."""
        selected_theme = self.theme_combobox.get()
        self.style.theme_use(selected_theme)
        self.current_theme_name = selected_theme
        self._save_theme_setting(selected_theme) # Salva o tema selecionado

    def _load_theme_setting(self):
        """Carrega o tema salvo do arquivo de configurações."""
        # Cria o arquivo de tema se ele não existir
        if not os.path.exists(self.theme_settings_path):
            try:
                with open(self.theme_settings_path, 'w', encoding='utf-8') as f:
                    f.write("cosmo") # Escreve o tema padrão inicial
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo de configurações do tema: {e}")
                return "cosmo" # Retorna o padrão em caso de erro na criação

        # Tenta ler o tema do arquivo
        try:
            with open(self.theme_settings_path, 'r', encoding='utf-8') as f:
                theme = f.read().strip()
                # Verifica se o tema lido é um tema válido do ttkbootstrap
                # self.style.theme_names() já está disponível neste ponto devido à inicialização prévia
                if theme in self.style.theme_names(): 
                    return theme
        except Exception as e:
            messagebox.showerror("Erro ao Carregar Tema", f"Não foi possível carregar o tema salvo: {e}\nO tema padrão será usado.")
        return "cosmo" # Retorna o tema padrão se não houver arquivo ou houver erro

    def _save_theme_setting(self, theme_name):
        """Salva o tema atual no arquivo de configurações."""
        try:
            with open(self.theme_settings_path, 'w', encoding='utf-8') as f:
                f.write(theme_name)
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Tema", f"Não foi possível salvar o tema: {e}\nPor favor, verifique as permissões de arquivo.")

    def on_search_key_release(self, event=None):
        """Atualiza a tabela quando o usuário digita na barra de pesquisa."""
        self.atualizar_tabela()

    def clear_search(self):
        """Limpa a barra de pesquisa e atualiza a tabela."""
        self.search_entry.delete(0, tk.END)
        self.atualizar_tabela()

    def atualizar_tabela(self):
        """
        Atualiza a Treeview com os dados mais recentes do banco de dados,
        incluindo uma linha de total e aplicando o filtro de pesquisa.
        """
        self.limpar_tabela_visual()
        
        search_term = self.search_entry.get().strip()
        vendas = self.db_manager.fetch_all_sales(search_term)
        
        total_geral_preco_final = 0.0

        for venda_db in vendas:
            venda_id = venda_db[0]
            
            preco_formatado = f"{venda_db[4]:.2f}".replace(".",",") if venda_db[4] is not None else ""
            preco_final_formatado = f"{venda_db[6]:.2f}".replace(".",",") if venda_db[6] is not None else ""

            # A ordem dos valores na tupla deve corresponder à ordem das colunas em self.tree.columns
            # (data_hora, nome_cliente, nome_produto, preco, tipo_pagamento, preco_final, nome_vendedor)
            dados_para_treeview = (
                venda_db[1], # data_hora
                venda_db[2], # nome_cliente
                venda_db[3], # nome_produto
                preco_formatado, # preco
                venda_db[5], # tipo_pagamento
                preco_final_formatado, # preco_final
                venda_db[7]  # nome_vendedor
            )
            self.tree.insert('', tk.END, iid=venda_id, values=dados_para_treeview)
            total_geral_preco_final += (venda_db[6] if venda_db[6] is not None else 0.0)

        total_final_formatado = f"{total_geral_preco_final:.2f}".replace(".", ",")
        # Os valores para a linha total também devem corresponder à ordem das colunas
        self.tree.insert('', tk.END, iid="total_row_id", values=(
            "", # data_hora
            "", # nome_cliente
            "", # nome_produto
            "", # preco
            "TOTAL GERAL:", # tipo_pagamento (usado para rótulo)
            total_final_formatado, # preco_final
            "" # nome_vendedor
        ), tags=('total_row',))
        
        self.tree.tag_configure('total_row', background='#e0e0e0', font=('Arial', 10, 'bold'))

    def fechar_app(self):
        """Fecha a conexão com o banco de dados antes de fechar a aplicação."""
        self.db_manager.close_connection()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesApp(root)
    root.mainloop()



