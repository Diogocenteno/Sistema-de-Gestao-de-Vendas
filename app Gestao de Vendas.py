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
    """
    Gerencia a conexão e as operações com o banco de dados SQLite.
    Responsável por inicializar o DB, inserir, atualizar, deletar e buscar vendas.
    """
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._initialize_db()

    def _initialize_db(self):
        """
        Inicializa a conexão com o banco de dados e cria a tabela 'vendas' se não existir.
        Também adiciona a coluna 'quantidade' se ela ainda não existir na tabela 'vendas',
        garantindo compatibilidade com versões anteriores do banco de dados.
        Oculta o arquivo do banco de dados no Windows para evitar manipulação acidental.
        """
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            
            # 1. Cria a tabela 'vendas' se ela não existir
            self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS vendas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_cliente TEXT,
                nome_produto TEXT,
                quantidade INTEGER, 
                preco REAL,
                tipo_pagamento TEXT,
                preco_final REAL,
                nome_vendedor TEXT,
                data_hora TEXT
            )
            """)
            self.conn.commit()

            # 2. Verifica se a coluna 'quantidade' existe e a adiciona se não existir
            self.cursor.execute("PRAGMA table_info(vendas)")
            columns = [col[1] for col in self.cursor.fetchall()]
            if 'quantidade' not in columns:
                try:
                    self.cursor.execute("ALTER TABLE vendas ADD COLUMN quantidade INTEGER")
                    self.conn.commit()
                    print("Coluna 'quantidade' adicionada à tabela 'vendas'.")
                except sqlite3.Error as e:
                    # Pode ocorrer se a coluna já foi adicionada por outro processo ou erro
                    print(f"Erro ao adicionar coluna 'quantidade': {e}")
                    self.conn.rollback() # Reverte se houver erro na alteração

            # Ocultar o arquivo do banco de dados (específico para Windows)
            if os.name == 'nt':
                try:
                    if os.path.exists(self.db_path):
                        # Define o atributo de arquivo como HIDDEN (0x02)
                        ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x02)
                except Exception as e:
                    print(f"Não foi possível ocultar o arquivo do banco de dados: {e}")
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao inicializar o banco de dados: {e}\nVerifique as permissões de arquivo ou se o banco de dados está corrompido.")
            exit() # Sai da aplicação se o banco de dados não puder ser inicializado

    def insert_sale(self, data):
        """Insere uma nova venda no banco de dados."""
        try:
            self.cursor.execute("""
            INSERT INTO vendas (nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor, data_hora)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                data["nome_cliente"], data["nome_produto"], data["quantidade"], data["preco"], data["tipo_pagamento"],
                data["preco_final"], data["nome_vendedor"], data["data_hora"]
            ))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao inserir venda: {e}\nPor favor, tente novamente. Se o problema persistir, reinicie a aplicação.")
            self.conn.rollback() # Reverte a transação em caso de erro

    def update_sale(self, sale_id, data):
        """Atualiza uma venda existente no banco de dados."""
        try:
            self.cursor.execute("""
            UPDATE vendas
            SET nome_cliente=?, nome_produto=?, quantidade=?, preco=?, tipo_pagamento=?, preco_final=?, nome_vendedor=?, data_hora=?
            WHERE id=?
            """, (
                data["nome_cliente"], data["nome_produto"], data["quantidade"], data["preco"], data["tipo_pagamento"],
                data["preco_final"], data["nome_vendedor"], data["data_hora"], sale_id
            ))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao atualizar venda: {e}\nNão foi possível salvar as alterações. Tente novamente.")
            self.conn.rollback() # Reverte a transação em caso de erro

    def delete_sale(self, sale_id):
        """Exclui uma venda do banco de dados pelo ID."""
        try:
            self.cursor.execute("DELETE FROM vendas WHERE id=?", (sale_id,))
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao excluir venda: {e}\nNão foi possível remover o registro. Tente novamente.")
            self.conn.rollback() # Reverte a transação em caso de erro

    def fetch_all_sales(self, search_term=""):
        """
        Busca todas as vendas do banco de dados, opcionalmente filtrando por um termo de busca.
        O termo de busca é aplicado aos campos nome_cliente, nome_produto e nome_vendedor.
        Retorna uma lista de tuplas com os dados das vendas.
        """
        try:
            query = "SELECT id, data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas"
            params = []
            if search_term:
                search_pattern = f"%{search_term}%"
                query += " WHERE nome_cliente LIKE ? OR nome_produto LIKE ? OR nome_vendedor LIKE ?"
                params = [search_pattern, search_pattern, search_pattern]
            query += " ORDER BY id DESC" # Ordena as vendas da mais recente para a mais antiga
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao buscar vendas: {e}\nNão foi possível carregar os dados da tabela. Verifique a conexão com o banco de dados.")
            return []

    def fetch_sale_by_id(self, sale_id):
        """Busca uma venda específica pelo ID."""
        try:
            # Inclui 'quantidade' na consulta para garantir que todos os campos sejam carregados
            self.cursor.execute("SELECT id, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, nome_vendedor FROM vendas WHERE id=?", (sale_id,))
            return self.cursor.fetchone()
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco de Dados", f"Erro ao buscar venda por ID: {e}\nNão foi possível recuperar os detalhes da venda para edição.")
            return None

    def close_connection(self):
        """
        Fecha a conexão com o banco de dados.
        No Windows, remove o atributo de oculto do arquivo do banco de dados.
        """
        if self.conn:
            self.conn.close()
        if os.name == 'nt' and os.path.exists(self.db_path):
            try:
                # Remove o atributo de oculto (0x02) para reexibir o arquivo.
                # O atributo normal é 0x80 (FILE_ATTRIBUTE_NORMAL), mas 0x00 também remove o oculto.
                # Usaremos 0x80 para garantir que ele seja visível.
                ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x80) 
            except Exception as e:
                print(f"Não foi possível reexibir o arquivo do banco de dados: {e}")

# --- Nova Classe para o Caderno Virtual de Encomendas ---
class CadernoVirtual:
    """
    Representa um caderno virtual para gerenciar encomendas.
    Permite adicionar, editar, excluir encomendas, calcular o total e exportar para Excel.
    As encomendas são salvas em um arquivo de texto simples.
    """
    DATA_DELIMITER = "|"  # Delimitador para salvar os dados no arquivo de texto
    TOTAL_ROW_ID = "caderno_total_row" # ID único para a linha de total na Treeview

    def __init__(self, parent_root, theme_style, file_path, anotacoes_file_path, excel_export_folder):
        self.file_path = file_path # Caminho para o arquivo de texto das encomendas
        self.anotacoes_file_path = anotacoes_file_path # Caminho para o arquivo de anotações
        self.excel_export_folder = excel_export_folder # Caminho para a pasta de exportação do Excel
        
        self.caderno_window = tk.Toplevel(parent_root)
        self.caderno_window.title("Caderno de Encomendas")
        self.caderno_window.geometry("1100x750") # Redimensionado para melhor visualização
        self.caderno_window.transient(parent_root) # Faz a janela do caderno ser filha da principal
        self.caderno_window.grab_set() # Impede interação com a janela pai enquanto o caderno está aberto
        # Garante que o conteúdo seja salvo ao fechar a janela
        self.caderno_window.protocol("WM_DELETE_WINDOW", self._on_closing) 

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

        # Título da janela do caderno
        ttk.Label(frame_caderno, text="Suas Encomendas", font=('Arial', 16, 'bold'), style='Caderno.TLabel').pack(pady=(0, 10))

        # Frame para os campos de entrada de nova encomenda
        frame_input = ttk.LabelFrame(frame_caderno, text="Adicionar/Editar Encomenda", padding=10)
        frame_input.pack(fill=tk.X, pady=(0, 10))

        # Campos de entrada para dados da encomenda
        ttk.Label(frame_input, text="Nome:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.nome_entry = ttk.Entry(frame_input, width=30)
        self.nome_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(frame_input, text="Produto:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.produto_entry = ttk.Entry(frame_input, width=30)
        self.produto_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(frame_input, text="Quantidade:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.quantidade_entry = ttk.Entry(frame_input, width=30)
        self.quantidade_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        # ALTERADO: "Valor (R$)" para "Valor por unidade (R$)"
        ttk.Label(frame_input, text="Valor por unidade (R$):").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.valor_entry = ttk.Entry(frame_input, width=30)
        self.valor_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        # Novo campo para Data de Entrega
        ttk.Label(frame_input, text="Data de Entrega:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.data_entrega_entry = ttk.Entry(frame_input, width=30) # Campo para entrada manual
        self.data_entrega_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=2)

        frame_input.grid_columnconfigure(1, weight=1)

        # Botões de ação para encomenda (Adicionar, Atualizar, Cancelar Edição)
        frame_botoes_input = ttk.Frame(frame_input)
        frame_botoes_input.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew") 

        self.btn_add_encomenda = ttk.Button(frame_botoes_input, text="Adicionar Encomenda", command=self._add_encomenda, style='success.Caderno.TButton')
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

        self.btn_update_encomenda = ttk.Button(frame_botoes_input, text="Atualizar Encomenda", command=self._update_encomenda, style='primary.Caderno.TButton')
        self.btn_update_encomenda.pack_forget() # Esconde inicialmente, aparece ao selecionar para editar

        self.btn_cancel_edit = ttk.Button(frame_botoes_input, text="Cancelar Edição", command=self._reset_input_fields, style='secondary.Caderno.TButton')
        self.btn_cancel_edit.pack_forget() # Esconde inicialmente

        # Treeview para exibir as encomendas
        # ALTERADO: Adicionado "ID" como primeira coluna
        cols = ("ID", "Data e Hora", "Nome", "Produto", "Quantidade", "Valor por unidade", "Data Entrega", "Valor Total") # Colunas atualizadas
        self.tree = ttk.Treeview(frame_caderno, columns=cols, show="headings", height=10, style='Caderno.Treeview')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Configuração dos cabeçalhos das colunas
        self.tree.heading("ID", text="ID") # Novo cabeçalho para ID
        self.tree.heading("Data e Hora", text="Data e Hora") 
        self.tree.heading("Nome", text="Nome do Cliente")
        self.tree.heading("Produto", text="Produto Encomendado")
        self.tree.heading("Quantidade", text="Qtd.") 
        # ALTERADO: "Valor" para "Valor por unidade" no cabeçalho da Treeview
        self.tree.heading("Valor por unidade", text="Valor por unidade (R$)") 
        self.tree.heading("Data Entrega", text="Data Entrega") 
        self.tree.heading("Valor Total", text="Valor Total (R$)") 

        # Larguras das colunas ajustadas para melhor visualização
        self.tree.column("ID", width=50, anchor="center", stretch=tk.NO) # Largura para ID
        self.tree.column("Data e Hora", width=150, anchor="center") 
        self.tree.column("Nome", width=150, anchor="center")
        self.tree.column("Produto", width=180, anchor="center")
        self.tree.column("Quantidade", width=80, anchor="center") 
        self.tree.column("Valor por unidade", width=100, anchor="center") # Largura ajustada
        self.tree.column("Data Entrega", width=120, anchor="center") 
        self.tree.column("Valor Total", width=110, anchor="center")

        # Scrollbars para a Treeview
        scrollbar_y = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(self.tree, orient="horizontal", command=self.tree.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        # Evento de seleção na Treeview para carregar dados para edição
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.selected_item_iid = None # Para controlar o item selecionado para edição
        self.next_encomenda_id = 1 # Para gerar IDs sequenciais para novas encomendas

        # Botões de gestão do caderno e o novo botão "Anotações"
        frame_botoes_gestao = ttk.Frame(frame_caderno, style='Caderno.TFrame')
        frame_botoes_gestao.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(frame_botoes_gestao, text="Excluir Encomenda Selecionada", command=self._delete_encomenda, style='danger.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes_gestao, text="Limpar Caderno Completo", command=self._clear_content, style='warning.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        
        # Botão para abrir a janela de anotações
        ttk.Button(frame_botoes_gestao, text="Abrir Anotações", command=self._abrir_anotacoes, style='info.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        
        # Botões de exportação e abertura de Excel no caderno
        ttk.Button(frame_botoes_gestao, text="Exportar Encomendas para Excel", command=self._exportar_excel_caderno, style='primary.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes_gestao, text="Abrir Planilha Excel (Encomendas)", command=self._abrir_planilha_excel_caderno, style='info.Caderno.TButton').pack(side=tk.LEFT, expand=True, padx=5)

        ttk.Button(frame_botoes_gestao, text="Fechar Caderno", command=self._on_closing, style='secondary.Caderno.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

        self._load_content() # Carrega o conteúdo do arquivo ao iniciar
        self._calculate_total() # Calcula o total inicial das encomendas
        self._reset_input_fields() # Limpa os campos e preenche a data e hora para a próxima entrada
        self.caderno_window.mainloop() # Inicia o loop para a nova janela

    def _load_content(self):
        """
        Carrega o conteúdo do arquivo de texto das encomendas para a Treeview.
        Formato esperado por linha: ID|Data e Hora|Nome|Produto|Quantidade|Valor Unitário|Data Entrega
        Se o arquivo estiver no formato antigo (sem ID), um ID será gerado.
        """
        self.tree.delete(*self.tree.get_children()) # Limpa a Treeview antes de carregar
        self.next_encomenda_id = 1 # Reseta o contador de ID ao carregar
        
        lines_to_write_back = [] # Para reescrever o arquivo se o formato for atualizado
        file_modified = False

        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        parts = line.strip().split(self.DATA_DELIMITER)
                        
                        current_id = None
                        # Tenta ler o ID. Se falhar, assume formato antigo e gera um ID.
                        if len(parts) > 6: # Novo formato com ID (ID|...)
                            try:
                                current_id = int(parts[0])
                                # Remove o ID da lista de partes para processar os dados restantes
                                data_parts = parts[1:] 
                            except ValueError:
                                # Se o primeiro campo não é um ID, assume formato antigo
                                data_parts = parts
                                current_id = self.next_encomenda_id
                                file_modified = True # Marca para reescrever o arquivo
                        elif len(parts) == 6: # Formato antigo (sem ID)
                            data_parts = parts
                            current_id = self.next_encomenda_id
                            file_modified = True # Marca para reescrever o arquivo
                        else:
                            print(f"Linha mal formatada no arquivo de encomendas: {line.strip()}")
                            continue # Pula linhas mal formatadas

                        # Atualiza o próximo ID disponível
                        self.next_encomenda_id = max(self.next_encomenda_id, current_id + 1)

                        # Espera 6 partes de dados: Data e Hora, Nome, Produto, Quantidade, Valor Unitário, Data Entrega
                        if len(data_parts) == 6:
                            try:
                                data_hora_registro = data_parts[0]
                                nome = data_parts[1]
                                produto = data_parts[2]
                                quantidade = int(data_parts[3])
                                valor_unitario = float(data_parts[4].replace(",", "."))
                                data_entrega = data_parts[5] 
                                valor_total = quantidade * valor_unitario
                                
                                # Insere na Treeview com o ID como iid
                                self.tree.insert('', tk.END, iid=current_id, values=(
                                    current_id, data_hora_registro, nome, produto, quantidade, 
                                    f"{valor_unitario:.2f}".replace(".", ","), 
                                    data_entrega, 
                                    f"{valor_total:.2f}".replace(".", ",")
                                ))
                                # Adiciona a linha ao buffer para reescrita (garantindo o formato com ID)
                                lines_to_write_back.append(f"{current_id}{self.DATA_DELIMITER}{data_hora_registro}{self.DATA_DELIMITER}{nome}{self.DATA_DELIMITER}"
                                                           f"{produto}{self.DATA_DELIMITER}{quantidade}{self.DATA_DELIMITER}"
                                                           f"{f'{valor_unitario:.2f}'.replace('.', ',')}{self.DATA_DELIMITER}{data_entrega}\n")
                            except ValueError:
                                print(f"Erro de conversão de dados na linha (ID: {current_id}): {line.strip()}")
                                # Em caso de erro, ainda tenta adicionar a linha bruta para não perder dados
                                self.tree.insert('', tk.END, iid=current_id, values=(current_id, *data_parts, ""))
                                file_modified = True # Marca para reescrever o arquivo
                        else:
                            print(f"Linha com número incorreto de partes (ID: {current_id}): {line.strip()}")
                            # Adiciona a linha bruta para não perder dados
                            self.tree.insert('', tk.END, iid=current_id, values=(current_id, *data_parts, "", "", ""))
                            file_modified = True # Marca para reescrever o arquivo

            except Exception as e:
                messagebox.showerror("Erro ao Carregar Caderno", f"Não foi possível carregar o caderno de encomendas: {e}")
        else:
            # Se o arquivo não existe, tenta criá-lo
            try:
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.write("")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo do caderno: {e}")
        
        # Se o arquivo foi modificado (formato antigo detectado), reescreve-o
        if file_modified:
            try:
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.writelines(lines_to_write_back)
                messagebox.showinfo("Atualização de Arquivo", "O formato do arquivo de encomendas foi atualizado para incluir IDs. O arquivo foi salvo novamente.")
            except Exception as e:
                messagebox.showwarning("Aviso de Salvamento", f"Não foi possível reescrever o arquivo de encomendas com os novos IDs: {e}")

        self._calculate_total()


    def _save_content(self):
        """
        Salva o conteúdo da Treeview para o arquivo de texto das encomendas.
        A linha de total não é salva no arquivo.
        """
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                for item_id in self.tree.get_children():
                    if item_id == self.TOTAL_ROW_ID: # Não salva a linha de total
                        continue
                    values = self.tree.item(item_id, 'values')
                    # values agora tem 8 elementos: ID, Data e Hora, Nome, Produto, Quantidade, Valor Unit., Data Entrega, Valor Total
                    # Precisamos salvar ID, Data e Hora, Nome, Produto, Quantidade, Valor Unitário e Data Entrega (índices 0 a 6)
                    
                    # O ID já está no values[0], mas o iid do item da treeview é a fonte mais confiável para o ID
                    encomenda_id = item_id 
                    data_hora_registro = values[1]
                    nome = values[2]
                    produto = values[3]
                    quantidade = values[4]
                    # Garante que o valor unitário seja salvo com ponto como separador decimal para consistência
                    valor_unitario_salvar = values[5].replace(",", ".") if values[5] else "0.00"
                    data_entrega = values[6] 
                    
                    try:
                        float(valor_unitario_salvar) # Tenta converter para float para validar
                    except ValueError:
                        valor_unitario_salvar = "0.00" # Define como 0.00 se for inválido

                    # Escreve o ID junto com os outros dados
                    f.write(f"{encomenda_id}{self.DATA_DELIMITER}{data_hora_registro}{self.DATA_DELIMITER}{nome}{self.DATA_DELIMITER}{produto}{self.DATA_DELIMITER}"
                            f"{quantidade}{self.DATA_DELIMITER}{valor_unitario_salvar}{self.DATA_DELIMITER}{data_entrega}\n")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Caderno", f"Não foi possível salvar o caderno de encomendas: {e}")
        self._calculate_total() # Recalcula o total após salvar

    def _add_encomenda(self):
        """Adiciona uma nova encomenda à Treeview e salva no arquivo."""
        nome = self.nome_entry.get().strip()
        produto = self.produto_entry.get().strip()
        quantidade_str = self.quantidade_entry.get().strip()
        valor_unitario_str = self.valor_entry.get().strip()
        data_entrega = self.data_entrega_entry.get().strip() 
        data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S") # Obtém a data e hora atual do registro

        if not nome or not produto or not quantidade_str or not valor_unitario_str or not data_entrega:
            messagebox.showwarning("Atenção", "Todos os campos (Nome, Produto, Quantidade, Valor por unidade, Data de Entrega) são obrigatórios.")
            return

        try:
            quantidade = int(quantidade_str)
            if quantidade <= 0:
                messagebox.showwarning("Atenção", "A quantidade deve ser um número inteiro positivo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Quantidade inválida! Digite um número inteiro válido (ex: 1, 2, 3).")
            return

        try:
            valor_unitario = float(valor_unitario_str.replace(",", "."))
            if valor_unitario < 0:
                messagebox.showwarning("Atenção", "O valor não pode ser negativo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Valor inválido! Digite um número válido (ex: 10,50 ou 10.50).")
            return

        valor_total = quantidade * valor_unitario
        
        # Formata os valores para exibição na Treeview (com vírgula para decimal)
        valor_unitario_exibicao = f"{valor_unitario:.2f}".replace(".", ",")
        valor_total_exibicao = f"{valor_total:.2f}".replace(".", ",")
        
        # Gera um novo ID para a encomenda
        new_id = self.next_encomenda_id
        self.next_encomenda_id += 1 # Incrementa para o próximo ID

        # Insere a nova encomenda na Treeview, incluindo o ID
        self.tree.insert('', tk.END, iid=new_id, values=(
            new_id, data_hora_atual, nome, produto, quantidade, 
            valor_unitario_exibicao, data_entrega, valor_total_exibicao
        ))
        self._save_content() # Salva as alterações no arquivo
        self._reset_input_fields() # Limpa os campos após adicionar
        self._calculate_total() # Recalcula o total

    def _on_tree_select(self, event):
        """
        Carrega os dados da encomenda selecionada na Treeview para os campos de entrada
        para permitir a edição. Ajusta a visibilidade dos botões.
        """
        selected_items = self.tree.selection()
        if not selected_items:
            self.selected_item_iid = None
            self._reset_input_fields()
            return

        # Impede a seleção da linha de total para edição
        if selected_items[0] == self.TOTAL_ROW_ID:
            self.tree.selection_remove(selected_items[0])
            self.selected_item_iid = None
            self._reset_input_fields()
            return

        self.selected_item_iid = selected_items[0]
        values = self.tree.item(self.selected_item_iid, 'values')

        # Preenche os campos de entrada com os valores da encomenda selecionada
        # values agora tem 8 elementos: ID, Data e Hora, Nome, Produto, Quantidade, Valor Unit., Data Entrega, Valor Total
        self.nome_entry.delete(0, tk.END)
        self.nome_entry.insert(0, values[2]) # Nome (índice 2, pois o ID está no índice 0 e Data e Hora no 1)
        self.produto_entry.delete(0, tk.END)
        self.produto_entry.insert(0, values[3]) # Produto
        self.quantidade_entry.delete(0, tk.END) 
        self.quantidade_entry.insert(0, values[4]) # Quantidade
        self.valor_entry.delete(0, tk.END)
        self.valor_entry.insert(0, values[5]) # Valor Unitário
        self.data_entrega_entry.delete(0, tk.END) 
        self.data_entrega_entry.insert(0, values[6]) # Data de Entrega

        # Altera a visibilidade dos botões: esconde "Adicionar", mostra "Atualizar" e "Cancelar Edição"
        self.btn_add_encomenda.pack_forget()
        self.btn_update_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        self.btn_cancel_edit.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

    def _update_encomenda(self):
        """
        Atualiza a encomenda selecionada na Treeview com os novos dados dos campos de entrada
        e salva as alterações no arquivo.
        """
        if not self.selected_item_iid:
            messagebox.showwarning("Atenção", "Nenhuma encomenda selecionada para atualizar.")
            return

        # A data e hora de registro original é mantida a partir do item selecionado na treeview
        # O ID da encomenda é o self.selected_item_iid
        current_id = self.selected_item_iid
        data_hora_registro = self.tree.item(self.selected_item_iid, 'values')[1] # Data e Hora está no índice 1
        nome = self.nome_entry.get().strip()
        produto = self.produto_entry.get().strip()
        quantidade_str = self.quantidade_entry.get().strip()
        valor_unitario_str = self.valor_entry.get().strip()
        data_entrega = self.data_entrega_entry.get().strip() 

        # ALTERADO: Mensagem de aviso de campos obrigatórios
        if not nome or not produto or not quantidade_str or not valor_unitario_str or not data_entrega:
            messagebox.showwarning("Atenção", "Todos os campos (Nome, Produto, Quantidade, Valor por unidade, Data de Entrega) são obrigatórios.")
            return

        try:
            quantidade = int(quantidade_str)
            if quantidade <= 0:
                messagebox.showwarning("Atenção", "A quantidade deve ser um número inteiro positivo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Quantidade inválida! Digite um número inteiro válido.")
            return

        try:
            valor_unitario = float(valor_unitario_str.replace(",", "."))
            if valor_unitario < 0:
                messagebox.showwarning("Atenção", "O valor não pode ser negativo.")
                return
        except ValueError:
            messagebox.showwarning("Atenção", "Valor inválido! Digite um número válido (ex: 10,50 ou 10.50).")
            return
        
        valor_total = quantidade * valor_unitario
        valor_unitario_exibicao = f"{valor_unitario:.2f}".replace(".", ",")
        valor_total_exibicao = f"{valor_total:.2f}".replace(".", ",")

        # Atualiza o item na Treeview, incluindo o ID
        self.tree.item(self.selected_item_iid, values=(
            current_id, data_hora_registro, nome, produto, quantidade, 
            valor_unitario_exibicao, data_entrega, valor_total_exibicao
        ))
        self._save_content() # Salva as alterações no arquivo
        self._reset_input_fields() # Limpa os campos e redefine os botões
        messagebox.showinfo("Sucesso", "Encomenda atualizada com sucesso!")
        self._calculate_total() # Recalcula o total

    def _delete_encomenda(self):
        """Exclui a(s) encomenda(s) selecionada(s) da Treeview e salva no arquivo."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Atenção", "Selecione uma ou mais encomendas para excluir.")
            return

        if self.TOTAL_ROW_ID in selected_items:
            messagebox.showwarning("Atenção", "A linha de total não pode ser excluída. Por favor, desfaça a seleção da linha de total e tente novamente.")
            return

        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir {len(selected_items)} encomenda(s) selecionada(s)? Esta ação não pode ser desfeita."):
            try:
                for item_id in selected_items:
                    self.tree.delete(item_id) # Remove o item da Treeview
                self._save_content() # Salva o conteúdo atualizado no arquivo
                self._reset_input_fields() # Limpa os campos e redefine os botões
                messagebox.showinfo("Sucesso", f"{len(selected_items)} encomenda(s) excluída(s) com sucesso!")
                self._calculate_total() # Recalcula o total
            except Exception as e:
                messagebox.showerror("Erro ao Excluir", f"Ocorreu um erro ao excluir encomenda(s): {e}")

    def _reset_input_fields(self):
        """
        Limpa todos os campos de entrada do formulário de encomenda
        e redefine os botões para o estado de "Adicionar Encomenda".
        """
        self.nome_entry.delete(0, tk.END)
        self.produto_entry.delete(0, tk.END)
        self.quantidade_entry.delete(0, tk.END) 
        self.valor_entry.delete(0, tk.END)
        self.data_entrega_entry.delete(0, tk.END) 
        
        self.selected_item_iid = None # Reseta o item selecionado
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        self.btn_update_encomenda.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.tree.selection_remove(self.tree.selection()) # Desseleciona qualquer item na Treeview

    def _clear_content(self):
        """
        Limpa todo o conteúdo do caderno de encomendas (Treeview e arquivo)
        após uma confirmação do usuário.
        """
        if messagebox.askyesno("Confirmar Limpar", "Tem certeza que deseja limpar todo o conteúdo do caderno? Esta ação não pode ser desfeita."):
            self.tree.delete(*self.tree.get_children()) # Limpa a Treeview
            self._save_content() # Salva o arquivo vazio
            self._reset_input_fields()
            self.next_encomenda_id = 1 # Reseta o contador de ID ao limpar
            messagebox.showinfo("Caderno Limpo", "O caderno foi limpo.")
            self._calculate_total() # Recalcula o total (que será 0)

    def _calculate_total(self):
        """
        Calcula o total dos valores de todas as encomendas na Treeview
        e exibe esse total em uma linha especial na parte inferior da Treeview.
        """
        # Remove a linha de total existente, se houver, antes de recalcular
        if self.tree.exists(self.TOTAL_ROW_ID):
            self.tree.delete(self.TOTAL_ROW_ID)

        total_valor_geral = 0.0
        for item_id in self.tree.get_children():
            if item_id == self.TOTAL_ROW_ID: # Garante que a linha de total não seja incluída no cálculo
                continue
            
            values = self.tree.item(item_id, 'values')
            # values agora tem 8 elementos: ID, Data e Hora, Nome, Produto, Quantidade, Valor Unit., Data Entrega, Valor Total
            # O valor total está na coluna 7 (índice 7)
            if len(values) == 8: 
                try:
                    valor_total_str = values[7].replace(",", ".")
                    total_valor_geral += float(valor_total_str)
                except ValueError:
                    pass # Ignora valores não numéricos ou inválidos que podem estar na coluna de total

        total_formatado = f"{total_valor_geral:.2f}".replace(".", ",")
        # Insere a nova linha de total na Treeview
        # Ajustado para as novas colunas (adicionado um campo vazio para o ID)
        self.tree.insert('', tk.END, iid=self.TOTAL_ROW_ID, values=(
            "", "", "", "", "", "", "TOTAL GERAL:", total_formatado 
        ), tags=('total_row',))
        
        # Configura o estilo para a linha de total, usando a cor de fundo do tema
        total_row_bg_color = self.style.lookup('TFrame', 'background')
        self.style.configure('total_row.Caderno.Treeview', background=total_row_bg_color, font=('Arial', 10, 'bold'))
        self.tree.tag_configure('total_row', background=total_row_bg_color, font=('Arial', 10, 'bold'))


    def _abrir_anotacoes(self):
        """Abre a janela virtual para anotações."""
        AnotacoesVirtual(self.caderno_window, self.style, self.anotacoes_file_path)

    def _exportar_excel_caderno(self):
        """
        Exporta os dados das encomendas para uma nova aba ('Encomendas')
        no arquivo Excel existente. Se o arquivo não existir, ele é criado.
        Preserva a aba 'Vendas' se ela já existir no arquivo.
        """
        if not messagebox.askyesno("Confirmar Exportação", "Deseja realmente exportar as encomendas para o arquivo Excel?\nIsso criará ou atualizará a aba 'Encomendas'."):
            return

        excel_file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        
        # Carrega os dados existentes da aba 'Vendas' se o arquivo existir
        df_vendas_existente = pd.DataFrame()
        if os.path.exists(excel_file_path):
            try:
                df_vendas_existente = pd.read_excel(excel_file_path, sheet_name='Vendas')
            except ValueError:
                # A aba 'Vendas' pode não existir, o que é normal se for a primeira exportação de encomendas
                pass
            except Exception as e:
                messagebox.showwarning("Aviso de Leitura de Excel", f"Não foi possível ler a aba 'Vendas' do arquivo Excel: {e}\nContinuando com a exportação de encomendas.")

        try:
            data_to_export = []
            for item_id in self.tree.get_children():
                if item_id == self.TOTAL_ROW_ID:
                    continue # Não exporta a linha de total
                values = self.tree.item(item_id, 'values')
                # values agora tem 8 elementos: ID, Data e Hora, Nome, Produto, Quantidade, Valor Unit., Data Entrega, Valor Total
                if len(values) == 8: # Garante que a linha tem todos os dados esperados
                    try:
                        encomenda_id = values[0] # Pega o ID da encomenda
                        quantidade = int(values[4])
                        valor_unitario = float(values[5].replace(",", "."))
                        valor_total = float(values[7].replace(",", "."))
                    except ValueError:
                        # Em caso de erro de conversão, usa valores padrão para evitar quebrar a exportação
                        encomenda_id = "" # ID pode ser vazio se houver erro
                        quantidade = 0
                        valor_unitario = 0.0
                        valor_total = 0.0
                    
                    data_to_export.append({
                        "ID da Encomenda": encomenda_id, # Adicionado ID
                        "Data e Hora do Registro": values[1],
                        "Nome do Cliente": values[2],
                        "Produto Encomendado": values[3],
                        "Quantidade": quantidade,
                        "Valor por unidade (R$)": valor_unitario, 
                        "Data de Entrega": values[6],
                        "Valor Total da Encomenda (R$)": valor_total
                    })

            if not data_to_export:
                messagebox.showinfo("Informação", "Não há encomendas para exportar.")
                # Se não há encomendas, mas há dados de vendas existentes, ainda podemos reescrever o arquivo com apenas vendas
                if df_vendas_existente.empty:
                    return # Se não há nem vendas nem encomendas, não há nada para exportar

            df_encomendas = pd.DataFrame(data_to_export)

            # Adiciona a linha de total ao DataFrame de encomendas
            total_encomendas = df_encomendas["Valor Total da Encomenda (R$)"].sum()
            total_row_df_encomendas = pd.DataFrame([{
                "ID da Encomenda": "", # Vazio para a linha de total
                "Data e Hora do Registro": "",
                "Nome do Cliente": "",
                "Produto Encomendado": "",
                "Quantidade": float('nan'), # Usar NaN para colunas numéricas sem valor
                "Valor por unidade (R$)": float('nan'), 
                "Data de Entrega": "TOTAL GERAL:",
                "Valor Total da Encomenda (R$)": total_encomendas
            }])
            df_encomendas_final = pd.concat([df_encomendas, total_row_df_encomendas], ignore_index=True)

            os.makedirs(self.excel_export_folder, exist_ok=True) # Garante que a pasta de exportação exista
            
            with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                # Escreve a aba de Vendas (existente ou vazia)
                df_vendas_existente.to_excel(writer, sheet_name='Vendas', index=False)
                
                # Escreve a aba de Encomendas
                df_encomendas_final.to_excel(writer, sheet_name='Encomendas', index=False)

                workbook = writer.book
                
                # Formatação para a aba 'Vendas' (se houver dados)
                worksheet_vendas = writer.sheets['Vendas']
                money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
                # Aplica formatação de moeda às colunas de valor na aba 'Vendas'
                if not df_vendas_existente.empty:
                    worksheet_vendas.set_column('F:F', None, money_format) # Coluna 'Valor (R$)'
                    worksheet_vendas.set_column('H:H', None, money_format) # Coluna 'Preço Final (R$)'

                # Formatação para a aba 'Encomendas'
                worksheet_encomendas = writer.sheets['Encomendas']
                # Aplica formatação de moeda às colunas de valor na aba 'Encomendas'
                worksheet_encomendas.set_column('F:F', None, money_format) # Coluna 'Valor por unidade (R$)'
                worksheet_encomendas.set_column('H:H', None, money_format) # Coluna 'Valor Total da Encomenda (R$)'

                # Auto-ajuste das colunas para a aba 'Vendas'
                for i, col in enumerate(df_vendas_existente.columns):
                    max_len = max(len(str(col)), df_vendas_existente[col].astype(str).map(len).max())
                    worksheet_vendas.set_column(i, i, max_len + 2) # +2 para um pequeno padding

                # Auto-ajuste das colunas para a aba 'Encomendas'
                for i, col in enumerate(df_encomendas_final.columns):
                    max_len = max(len(str(col)), df_encomendas_final[col].astype(str).map(len).max())
                    worksheet_encomendas.set_column(i, i, max_len + 2) # +2 para um pequeno padding
                    
            messagebox.showinfo("Exportado", f"Encomendas exportadas com sucesso para a aba 'Encomendas' em:\n{excel_file_path}")
        except Exception as e:
            # Formata a mensagem de erro para evitar problemas com '%' em messagebox
            error_message = str(e).replace('%', '%%') 
            messagebox.showerror("Erro", f"Erro ao exportar encomendas para Excel: {error_message}\nCertifique-se de que o arquivo Excel não esteja aberto em outro programa e que a biblioteca 'openpyxl' esteja instalada (`pip install openpyxl`).")


    def _abrir_planilha_excel_caderno(self):
        """
        Abre o arquivo Excel exportado (o mesmo usado para vendas e encomendas), se existir.
        Tenta abrir com o aplicativo padrão do sistema operacional.
        """
        file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Atenção", "O arquivo Excel ainda não foi exportado.\nPor favor, exporte-o primeiro clicando no botão 'Exportar Encomendas para Excel' ou 'Exportar para Excel' na tela principal.")
            return
        
        try:
            if os.name == 'nt': # Windows
                subprocess.Popen(['start', file_path], shell=True) # Use 'start' para abrir com o programa padrão
            elif os.uname().sysname == 'Darwin': # macOS
                subprocess.Popen(['open', file_path])
            else: # Linux/Outros Unix-like
                subprocess.Popen(['xdg-open', file_path])
        except FileNotFoundError:
            messagebox.showerror("Erro", "Nenhum aplicativo associado encontrado para abrir arquivos .xlsx.\nCertifique-se de ter o Microsoft Excel ou um programa compatível instalado.")
        except Exception as e:
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir a planilha: {error_message}\nVerifique se o arquivo Excel não está corrompido ou em uso por outro programa.")


    def _on_closing(self):
        """
        Salva o conteúdo do caderno de encomendas antes de fechar a janela.
        Libera o foco da janela para a janela pai.
        """
        self._save_content() # Garante que as últimas alterações sejam salvas
        self.caderno_window.destroy()
        self.caderno_window.grab_release() # Libera o foco para a janela pai


# --- NOVA CLASSE PARA AS ANOTAÇÕES VIRTUAIS ---
class AnotacoesVirtual:
    """
    Representa uma janela de anotações virtuais com um campo de texto e linhas de caderno.
    As anotações são salvas em um arquivo de texto simples.
    """
    def __init__(self, parent_root, theme_style, file_path):
        self.file_path = file_path # Caminho para o arquivo de texto das anotações
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

        # Frame para conter o canvas e o campo de texto
        frame_texto = ttk.Frame(frame_anotacoes)
        frame_texto.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Canvas para desenhar as linhas do caderno
        self.canvas = tk.Canvas(frame_texto, 
                                 bg=self.style.lookup('TEntry', 'fieldbackground'),
                                 highlightthickness=0) # Remove a borda padrão do canvas
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar vertical para o canvas
        scrollbar_y = ttk.Scrollbar(frame_texto, orient="vertical", command=self.canvas.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Campo de texto que será colocado DENTRO do canvas
        self.text_area = tk.Text(self.canvas, wrap=tk.WORD, font=('Courier New', 12),
                                 bg=self.style.lookup('TEntry', 'fieldbackground'),
                                 fg=self.style.lookup('TEntry', 'foreground'),
                                 insertbackground=self.style.lookup('TEntry', 'foreground'),
                                 relief=tk.FLAT, bd=2) # Sem borda para se integrar melhor ao canvas
        
        # Coloca o widget de texto no canvas na posição (0,0)
        # O text_window é o ID do objeto de janela dentro do canvas
        self.text_window = self.canvas.create_window((0, 0), window=self.text_area, anchor="nw")
        
        # Configura o scroll do canvas para controlar o scroll do text_area
        self.text_area.configure(yscrollcommand=scrollbar_y.set)

        # Configura eventos para redimensionamento e liberação de teclas
        # Quando o text_area é redimensionado ou o conteúdo muda, ajusta o canvas
        self.text_area.bind("<Configure>", self._ajustar_canvas)
        self.text_area.bind("<KeyRelease>", self._ajustar_canvas) 
        # Quando o canvas é redimensionado, ajusta o text_area e redesenha as linhas
        self.canvas.bind("<Configure>", self._ajustar_texto_e_linhas) 
        
        self._desenhar_linhas() # Desenha as linhas iniciais

        # Botões de ação para as anotações
        frame_botoes = ttk.Frame(frame_anotacoes, style='Anotacoes.TFrame')
        frame_botoes.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(frame_botoes, text="Salvar Anotações", command=self._save_content, style='info.Anotacoes.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Limpar Anotações", command=self._clear_content, style='danger.Anotacoes.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Fechar Anotações", command=self._on_closing, style='secondary.Anotacoes.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

        self._load_content() # Carrega o conteúdo das anotações ao iniciar
        self.anotacoes_window.mainloop()

    def _scroll_both(self, *args):
        """
        Método de rolagem que sincroniza a rolagem do canvas e da área de texto.
        (Atualmente não está diretamente ligado a um scrollbar, mas útil para depuração ou futuras expansões)
        """
        self.canvas.yview(*args)
        self.text_area.yview(*args)
        self._desenhar_linhas() # Redesenha as linhas ao rolar

    def _ajustar_canvas(self, event=None):
        """
        Ajusta a região de rolagem do canvas com base no conteúdo do text_area.
        Isso permite que o scrollbar do canvas funcione corretamente com o conteúdo do texto.
        Redesenha as linhas após o ajuste.
        """
        try:
            # Obtém as coordenadas da última linha do texto para determinar a altura total
            # O índice 'end-1c' representa o final do texto, excluindo o último caractere de nova linha
            total_text_height = self.text_area.winfo_height()
            if self.text_area.compare("end-1c", ">", "1.0"): # Se há algum texto
                last_line_info = self.text_area.dlineinfo("end-1c")
                if last_line_info:
                    # last_line_info é uma tupla (x, y, width, height, baseline).
                    # A altura total do conteúdo é a coordenada y da última linha + sua altura.
                    total_text_height = last_line_info[1] + last_line_info[3]
            
            # Garante que a altura mínima seja a altura visível do canvas
            total_text_height = max(total_text_height, self.canvas.winfo_height())
            
        except tk.TclError: 
            # Em caso de erro (ex: text_area ainda não totalmente renderizado), usa a altura do canvas
            total_text_height = self.canvas.winfo_height()

        # Configura a região de rolagem do canvas. O último valor é a altura total do conteúdo.
        self.canvas.config(scrollregion=(0, 0, self.canvas.winfo_width(), total_text_height))
        self._desenhar_linhas() # Redesenha as linhas para corresponder à nova altura

    def _ajustar_texto_e_linhas(self, event=None):
        """
        Ajusta o tamanho do widget de texto dentro do canvas e redesenha as linhas
        sempre que o canvas é redimensionado.
        """
        # Ajusta a largura e altura do widget de texto para preencher o canvas
        # O text_window é o ID do objeto de janela dentro do canvas
        self.canvas.itemconfig(self.text_window, width=self.canvas.winfo_width())
        self.canvas.itemconfig(self.text_window, height=self.canvas.winfo_height())
        self._desenhar_linhas() # Redesenha as linhas

    def _desenhar_linhas(self, event=None):
        """
        Desenha linhas horizontais no canvas para simular as linhas de um caderno.
        As linhas são desenhadas com base na altura da linha de texto e na região de rolagem do canvas.
        """
        self.canvas.delete("linha") # Remove todas as linhas existentes antes de redesenhar
        
        try:
            # Obtém a altura de uma linha de texto para espaçar as linhas do caderno
            line_height_info = self.text_area.dlineinfo("1.0")
            if not line_height_info:
                return # Não há informações de linha, não desenha
            line_height = line_height_info[3] # A altura da linha está no índice 3
            if line_height <= 0:
                return # Altura inválida
        except tk.TclError:
            return # Erro ao obter informações da linha

        # Determina a altura total desenhável do canvas
        # A scrollregion do canvas é uma tupla (x1, y1, x2, y2)
        scroll_region_coords = self.canvas.bbox("all") # Obtém o bounding box de todos os itens no canvas
        if scroll_region_coords:
            total_drawable_height = scroll_region_coords[3] # y2 do bounding box
        else:
            total_drawable_height = self.canvas.winfo_height() # Fallback para a altura visível

        # Desenha as linhas, estendendo um pouco além da altura visível para cobrir a rolagem
        # Começa a desenhar as linhas a partir da parte superior visível do canvas
        canvas_y_offset = self.canvas.canvasy(0) # Posição Y da parte superior visível do canvas
        
        # Calcula o número de linhas necessárias para cobrir a área visível e um pouco mais para rolagem
        num_lines = int((total_drawable_height + self.canvas.winfo_height()) / line_height) + 2

        for i in range(num_lines):
            y = i * line_height
            # Desenha a linha apenas se estiver dentro da área de rolagem do canvas
            if y >= canvas_y_offset - line_height and y <= canvas_y_offset + self.canvas.winfo_height() + line_height:
                self.canvas.create_line(0, y, self.canvas.winfo_width(), y, 
                                        fill="#cccccc", tags="linha") # Cor cinza claro para as linhas

    def _load_content(self):
        """Carrega o conteúdo do arquivo de anotações para o campo de texto."""
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.text_area.delete(1.0, tk.END) # Limpa o campo de texto
                    self.text_area.insert(1.0, content) # Insere o conteúdo do arquivo
            except Exception as e:
                messagebox.showerror("Erro ao Carregar Anotações", f"Não foi possível carregar as anotações: {e}")
        else:
            # Se o arquivo não existe, tenta criá-lo
            try:
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.write("")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo de anotações: {e}")
        self._ajustar_canvas() # Ajusta o canvas após carregar o conteúdo

    def _save_content(self):
        """Salva o conteúdo do campo de texto das anotações para o arquivo."""
        try:
            content = self.text_area.get(1.0, tk.END).strip() # Obtém o texto e remove espaços em branco extras
            with open(self.file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Salvo", "Anotações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Anotações", f"Não foi possível salvar as anotações: {e}")

    def _clear_content(self):
        """Limpa todo o conteúdo das anotações (campo de texto e arquivo) após confirmação."""
        if messagebox.askyesno("Confirmar Limpar", "Tem certeza que deseja limpar todo o conteúdo das anotações? Esta ação não pode ser desfeita."):
            self.text_area.delete(1.0, tk.END) # Limpa o campo de texto
            self._save_content() # Salva o arquivo vazio
            messagebox.showinfo("Anotações Limpas", "As anotações foram limpas.")
            self._ajustar_canvas() # Ajusta o canvas após limpar

    def _on_closing(self):
        """
        Salva o conteúdo das anotações antes de fechar a janela.
        Libera o foco da janela para a janela pai.
        """
        self._save_content()
        self.anotacoes_window.destroy()
        self.anotacoes_window.grab_release() # Libera o foco para a janela pai

# --- Classe Principal da Aplicação ---
class SalesApp:
    """
    Classe principal da aplicação de gestão de vendas.
    Gerencia a interface do usuário, interage com o banco de dados e abre as janelas de caderno/anotações.
    """
    TOTAL_ROW_ID = "total_row_id" # ID único para a linha de total na Treeview principal

    def __init__(self, root):
        self.root = root
        
        # Define os caminhos para os arquivos de dados
        app_dir = os.path.dirname(os.path.abspath(__file__))
        internal_data_dir = os.path.join(app_dir, "_internal")
        
        self.db_path = os.path.join(internal_data_dir, "vendas_padaria.db")
        self.caderno_path = os.path.join(internal_data_dir, "encomendas_caderno.txt")
        self.anotacoes_path = os.path.join(internal_data_dir, "anotacoes.txt")
        self.theme_settings_path = os.path.join(internal_data_dir, "theme_setting.txt")

        # Tenta criar o diretório '_internal'. Se falhar (ex: permissões), usa a pasta Documentos do usuário.
        try:
            os.makedirs(internal_data_dir, exist_ok=True)
            if not os.access(internal_data_dir, os.W_OK):
                raise OSError("Diretório '_internal' não é gravável.")
        except OSError:
            self.documents_path = os.path.join(os.path.expanduser("~"), "Documents")
            self.db_path = os.path.join(self.documents_path, "vendas_padaria.db")
            self.caderno_path = os.path.join(self.documents_path, "encomendas_caderno.txt")
            self.anotacoes_path = os.path.join(self.documents_path, "anotacoes.txt")
            self.theme_settings_path = os.path.join(self.documents_path, "theme_setting.txt")
            os.makedirs(os.path.dirname(self.db_path), exist_ok=True) # Garante que a pasta Documentos exista
            messagebox.showwarning("Aviso de Caminho de Dados",
                                   f"Não foi possível criar ou acessar a pasta '_internal' no diretório do aplicativo.\n"
                                   f"O banco de dados, caderno e anotações serão salvos em: {self.documents_path}")

        # Pasta para exportação de arquivos Excel (sempre na pasta Documentos do usuário)
        self.excel_export_folder = os.path.join(os.path.expanduser("~"), "Documents", "Vendas_Padaria_Exportadas")
        os.makedirs(self.excel_export_folder, exist_ok=True) # Garante que a pasta de exportação exista

        # Configuração do tema ttkbootstrap
        self.style = Style(theme="cosmo") # Tema padrão inicial
        self.current_theme_name = self._load_theme_setting() # Carrega o tema salvo
        self.style.theme_use(self.current_theme_name) # Aplica o tema
        
        self.style.configure("TButton", padding=(10, 5)) # Estilo padrão para botões

        self.root.title("Sistema de Cadastro de Vendas Padaria do Jeff e Pri")
        self.root.geometry("1050x800") # Tamanho inicial da janela principal

        self.db_manager = DatabaseManager(self.db_path) # Inicializa o gerenciador do banco de dados
        self.id_venda_em_edicao = None # Armazena o ID da venda sendo editada

        self._create_widgets() # Cria todos os elementos da interface
        self.atualizar_tabela() # Carrega os dados na tabela ao iniciar
        self.limpar_campos_e_resetar_edicao() # Limpa os campos do formulário

        # Define a função a ser chamada ao fechar a janela principal
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_app)

    def _create_widgets(self):
        """Cria e organiza todos os widgets da interface gráfica da aplicação principal."""
        title_label = ttk.Label(self.root, text="SISTEMA DE GESTÃO DE VENDAS", font=("Arial", 36, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 15))

        # Frame para seleção de tema
        frame_theme_selector = ttk.Frame(self.root, padding=(10, 5))
        frame_theme_selector.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(5, 0))

        ttk.Label(frame_theme_selector, text="Escolher Tema:").pack(side=tk.RIGHT, padx=(0, 5))
        self.available_themes = sorted(list(self.style.theme_names())) # Obtém todos os temas disponíveis
        self.theme_combobox = ttk.Combobox(frame_theme_selector, values=self.available_themes, state="readonly", width=20)
        self.theme_combobox.set(self.current_theme_name) # Define o tema atual como selecionado
        self.theme_combobox.pack(side=tk.RIGHT)
        self.theme_combobox.bind("<<ComboboxSelected>>", self.change_theme) # Evento para mudar o tema

        # Frame para o formulário de registro de venda
        frame_formulario = ttk.LabelFrame(self.root, text="Registo de Venda Padaria do Jeff e Pri", padding=(10,10))
        frame_formulario.grid(row=2, column=0, padx=10, pady=(5, 10), sticky="ew", columnspan=3)

        # Definição dos campos do formulário
        labels_info = [
            ("nome_cliente", "Nome do Cliente:"),
            ("nome_produto", "Nome do Produto:"),
            ("quantidade", "Quantidade:"), 
            # ALTERADO: "Valor (R$)" para "Valor por unidade (R$)"
            ("preco", "Valor por unidade (R$):"), 
            ("tipo_pagamento", "Tipo de Pagamento:"),
            ("nome_vendedor", "Nome do Vendedor:"),
        ]
        self.campos = {} # Dicionário para armazenar as referências aos widgets de entrada

        for i, (chave, texto_label) in enumerate(labels_info):
            label = ttk.Label(frame_formulario, text=texto_label)
            label.grid(row=i, column=0, sticky="w", padx=5, pady=5)
            
            if chave == "tipo_pagamento":
                # Combobox para o tipo de pagamento
                combo_pagamento = ttk.Combobox(frame_formulario, values=["Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Pix"], state="readonly", width=28)
                combo_pagamento.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.campos[chave] = combo_pagamento
            else:
                # Campo de entrada de texto genérico
                entrada = ttk.Entry(frame_formulario, width=30)
                entrada.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.campos[chave] = entrada

        frame_formulario.grid_columnconfigure(1, weight=1) # Faz a coluna de entrada expandir

        # Botões de ação do formulário (Salvar/Atualizar, Cancelar Edição)
        frame_botoes_formulario = ttk.Frame(frame_formulario)
        frame_botoes_formulario.grid(row=len(labels_info), column=0, columnspan=2, pady=10, sticky="ew")

        self.btn_salvar = ttk.Button(frame_botoes_formulario, text="Registar Venda", command=self.salvar_dados, style="success.TButton")
        self.btn_salvar.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.btn_cancelar_edicao = ttk.Button(frame_botoes_formulario, text="Cancelar Edição", command=self.limpar_campos_e_resetar_edicao, style="secondary.TButton")
        self.btn_cancelar_edicao.pack_forget() # Esconde inicialmente

        # Frame para botões de gestão de vendas e exportação
        frame_botoes_gestao = ttk.LabelFrame(self.root, text="Gestão de Vendas e Exportação", padding=(10,10))
        frame_botoes_gestao.grid(row=3, column=0, padx=10, pady=5, sticky="ew", columnspan=3)

        ttk.Button(frame_botoes_gestao, text="Carregar para Editar", command=self.carregar_para_edicao, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Excluir Venda Selecionada", command=self.excluir_venda_selecionada, style="danger.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Exportar para Excel", command=self.exportar_excel, style="primary.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Abrir Planilha Excel", command=self.abrir_planilha_excel, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Limpar Tabela (Visual)", command=self.limpar_tabela_visual, style="warning.TButton").pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Button(frame_botoes_gestao, text="Abrir Calculadora", command=self.abrir_calculadora, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        # Botão para abrir o caderno de encomendas
        ttk.Button(frame_botoes_gestao, text="Encomendas", command=self.abrir_caderno_encomendas, style="info.TButton").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # Frame para a tabela de vendas registradas
        frame_tabela = ttk.LabelFrame(self.root, text="Vendas Registradas", padding=(10,10))
        frame_tabela.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        # Configuração de expansão para a janela principal e a tabela
        self.root.grid_rowconfigure(4, weight=1) # Faz a linha da tabela expandir verticalmente
        self.root.grid_columnconfigure(0, weight=1) # Faz as colunas expandir horizontalmente
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

        frame_tabela.grid_rowconfigure(1, weight=1) # Faz a linha da Treeview dentro do frame expandir
        frame_tabela.grid_columnconfigure(0, weight=1) # Faz a coluna da Treeview dentro do frame expandir

        # Frame para a barra de pesquisa
        frame_pesquisa = ttk.Frame(frame_tabela, padding=(0, 5))
        frame_pesquisa.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        ttk.Label(frame_pesquisa, text="Pesquisar:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(frame_pesquisa, width=50)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.search_entry.bind("<KeyRelease>", self.on_search_key_release) # Atualiza a tabela ao digitar

        ttk.Button(frame_pesquisa, text="Limpar Pesquisa", command=self.clear_search, style="secondary.TButton").pack(side=tk.LEFT, padx=(5, 0))

        # Treeview para exibir as vendas
        # Colunas visíveis atualizadas para incluir 'quantidade'
        colunas_visiveis = ("data_hora", "nome_cliente", "nome_produto", "quantidade", "preco", "tipo_pagamento", "preco_final", "nome_vendedor")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas_visiveis, show="headings", height=10, selectmode="extended") 
        self.tree.grid(row=1, column=0, sticky="nsew")

        # Scrollbars para a Treeview
        scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree.xview)
        scrollbar_x.grid(row=2, column=0, sticky="ew", columnspan=2)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        # Configuração dos cabeçalhos das colunas da Treeview
        self.tree.heading("data_hora", text="Data e Hora")
        self.tree.heading("nome_cliente", text="Cliente")
        self.tree.heading("nome_produto", text="Produto")
        self.tree.heading("quantidade", text="Qtd.") 
        # ALTERADO: "Valor (R$)" para "Valor por unidade (R$)"
        self.tree.heading("preco", text="Valor por unidade (R$)") 
        self.tree.heading("tipo_pagamento", text="Pagamento")
        self.tree.heading("preco_final", text="Preço Final (R$)")
        self.tree.heading("nome_vendedor", text="Vendedor")

        # Configuração das larguras das colunas
        self.tree.column("data_hora", width=140, anchor="center", stretch=tk.NO)
        self.tree.column("nome_cliente", width=160, anchor="center", stretch=tk.YES)
        self.tree.column("nome_produto", width=160, anchor="center", stretch=tk.YES)
        self.tree.column("quantidade", width=70, anchor="center", stretch=tk.NO) 
        self.tree.column("preco", width=120, anchor="center", stretch=tk.NO) # Largura ajustada
        self.tree.column("tipo_pagamento", width=120, anchor="center", stretch=tk.NO)
        self.tree.column("preco_final", width=100, anchor="center", stretch=tk.NO)
        self.tree.column("nome_vendedor", width=120, anchor="center", stretch=tk.YES)

    def abrir_calculadora(self):
        """
        Abre a calculadora padrão do sistema operacional.
        Compatível com Windows, macOS e Linux.
        """
        try:
            if os.name == 'nt': # Windows
                subprocess.Popen(['calc.exe'])
            elif os.uname().sysname == 'Darwin': # macOS
                subprocess.Popen(['open', '-a', 'Calculator'])
            else: # Linux/Outros Unix-like
                # Tenta diferentes comandos para calculadoras comuns no Linux
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
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir a calculadora: {error_message}\nVerifique se o aplicativo da calculadora está funcionando corretamente.")

    def abrir_caderno_encomendas(self):
        """Abre a janela do caderno virtual para anotações de encomendas."""
        # Passa o caminho da pasta de exportação do Excel para a classe CadernoVirtual
        CadernoVirtual(self.root, self.style, self.caderno_path, self.anotacoes_path, self.excel_export_folder)

    def abrir_planilha_excel(self):
        """
        Abre o arquivo Excel exportado (o mesmo usado para vendas e encomendas), se existir.
        Tenta abrir com o aplicativo padrão do sistema operacional.
        """
        file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Atenção", "O arquivo Excel ainda não foi exportado.\nPor favor, exporte-o primeiro clicando no botão 'Exportar para Excel'.")
            return
        
        try:
            if os.name == 'nt': # Windows
                subprocess.Popen(['start', file_path], shell=True)
            elif os.uname().sysname == 'Darwin': # macOS
                subprocess.Popen(['open', file_path])
            else: # Linux/Outros Unix-like
                subprocess.Popen(['xdg-open', file_path])
        except FileNotFoundError:
            messagebox.showerror("Erro", "Nenhum aplicativo associado encontrado para abrir arquivos .xlsx.\nCertifique-se de ter o Microsoft Excel ou um programa compatível instalado.")
        except Exception as e:
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir a planilha: {error_message}\nVerifique se o arquivo Excel não está corrompido ou em uso por outro programa.")

    def salvar_dados(self):
        """
        Salva ou atualiza os dados de uma venda no banco de dados.
        Realiza validação de campos obrigatórios, formata o preço e calcula o preço final.
        """
        try:
            dados = {
                "nome_cliente": self.campos["nome_cliente"].get().strip(),
                "nome_produto": self.campos["nome_produto"].get().strip(),
                "quantidade": self.campos["quantidade"].get().strip(), 
                "preco": self.campos["preco"].get().strip(),
                "tipo_pagamento": self.campos["tipo_pagamento"].get().strip(),
                "nome_vendedor": self.campos["nome_vendedor"].get().strip(),
            }

            # --- Validação de campos obrigatórios ---
            if not dados["nome_cliente"]:
                messagebox.showwarning("Atenção", "O campo 'Nome do Cliente' é obrigatório e não pode estar vazio.")
                self.campos["nome_cliente"].focus_set()
                return
            if not dados["nome_produto"]:
                messagebox.showwarning("Atenção", "O campo 'Nome do Produto' é obrigatório e não pode estar vazio.")
                self.campos["nome_produto"].focus_set()
                return
            if not dados["quantidade"]:
                messagebox.showwarning("Atenção", "O campo 'Quantidade' é obrigatório e não pode estar vazio.")
                self.campos["quantidade"].focus_set()
                return
            # ALTERADO: Mensagem de aviso de campos obrigatórios
            if not dados["preco"]:
                messagebox.showwarning("Atenção", "O campo 'Valor por unidade' é obrigatório e não pode estar vazio.")
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

            # --- Validação e conversão da quantidade ---
            try:
                quantidade = int(dados["quantidade"])
                if quantidade <= 0:
                    messagebox.showwarning("Atenção", "A quantidade deve ser um número inteiro positivo.")
                    self.campos["quantidade"].focus_set()
                    return
                dados["quantidade"] = quantidade
            except ValueError:
                messagebox.showwarning("Atenção", "Quantidade inválida! Digite um número inteiro válido (ex: 1, 2, 3).")
                self.campos["quantidade"].focus_set()
                return

            # --- Validação e conversão do preço ---
            try:
                preco_str = dados["preco"].replace(",", ".") # Substitui vírgula por ponto para conversão
                preco_unitario = float(preco_str)
                if preco_unitario < 0:
                    messagebox.showwarning("Atenção", "O valor não pode ser um valor negativo.")
                    self.campos["preco"].focus_set()
                    return
                dados["preco"] = preco_unitario
            except ValueError:
                # ALTERADO: Mensagem de aviso de valor inválido
                messagebox.showwarning("Atenção", "Valor por unidade inválido! Digite um número válido (ex: 10,50 ou 10.50).\nVerifique se não há letras ou múltiplos pontos/vírgulas.")
                self.campos["preco"].focus_set()
                return

            preco_final = dados["quantidade"] * dados["preco"] # Calcula o preço final
            data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            dados["preco_final"] = preco_final
            dados["data_hora"] = data_hora_atual

            if self.id_venda_em_edicao is not None:
                # Se há um ID de venda em edição, atualiza a venda existente
                self.db_manager.update_sale(self.id_venda_em_edicao, dados)
                messagebox.showinfo("Sucesso", "Venda atualizada com sucesso!")
            else:
                # Caso contrário, insere uma nova venda
                self.db_manager.insert_sale(dados)
                messagebox.showinfo("Sucesso", "Venda registrada com sucesso!")

            self.limpar_campos_e_resetar_edicao() # Limpa o formulário e reseta o estado
            self.atualizar_tabela() # Atualiza a tabela de vendas

        except Exception as e:
            error_message = str(e).replace('%', '%%') # Formata a mensagem de erro
            messagebox.showerror("Erro", f"Erro ao salvar dados: {error_message}\nOcorreu um problema inesperado ao tentar registrar/atualizar a venda. Verifique os dados inseridos.")


    def limpar_campos_e_resetar_edicao(self):
        """
        Limpa todos os campos do formulário de registro de venda
        e redefine o estado dos botões para "Registar Venda".
        """
        self.id_venda_em_edicao = None # Reseta o ID da venda em edição
        for chave in self.campos:
            if isinstance(self.campos[chave], ttk.Combobox):
                self.campos[chave].set("") # Limpa Combobox
            else:
                if self.campos[chave]: # Verifica se o widget existe
                    self.campos[chave].delete(0, tk.END) # Limpa Entry

        self.btn_salvar.config(text="Registar Venda", style="success.TButton")
        self.btn_cancelar_edicao.pack_forget() # Esconde o botão de cancelar edição

    def carregar_para_edicao(self):
        """
        Carrega os dados da venda selecionada na tabela para os campos de edição do formulário.
        Permite a edição de apenas uma venda por vez.
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

        if item_id_selecionado == self.TOTAL_ROW_ID:
            messagebox.showwarning("Atenção", "A linha de total não pode ser editada.")
            return

        try:
            # O iid da Treeview corresponde ao ID da venda no banco de dados
            venda_id_db = int(item_id_selecionado) 
            # A ordem dos campos retornados por fetch_sale_by_id é:
            # id, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, nome_vendedor
            venda = self.db_manager.fetch_sale_by_id(venda_id_db)

            if venda:
                self.limpar_campos_e_resetar_edicao() # Limpa os campos antes de preencher

                self.campos["nome_cliente"].insert(0, venda[1])
                self.campos["nome_produto"].insert(0, venda[2])
                self.campos["quantidade"].insert(0, str(venda[3])) # Carrega a quantidade
                self.campos["preco"].insert(0, str(venda[4]).replace(".", ",")) # Preço unitário, formata para vírgula
                self.campos["tipo_pagamento"].set(venda[5])
                self.campos["nome_vendedor"].insert(0, venda[6])
                
                self.id_venda_em_edicao = venda_id_db # Define o ID da venda em edição
                self.btn_salvar.config(text="Atualizar Venda", style="primary.TButton") # Altera o texto do botão
                
                self.btn_cancelar_edicao.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X) # Mostra o botão de cancelar

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
        """
        Exclui a(s) venda(s) selecionada(s) da tabela e do banco de dados.
        Permite a exclusão de múltiplas vendas.
        """
        selecionado = self.tree.selection()
        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione uma ou mais vendas na tabela para excluir.")
            return

        if self.TOTAL_ROW_ID in selecionado:
            messagebox.showwarning("Atenção", "A linha de total não pode ser excluída.\nPor favor, desfaça a seleção da linha de total e tente novamente.")
            return

        if len(selecionado) > 1:
            confirmation_message = f"Tem certeza que deseja excluir as {len(selecionado)} vendas selecionadas? Esta ação não pode ser desfeita."
        else:
            confirmation_message = "Tem certeza que deseja excluir a venda selecionada? Esta ação não pode ser desfeita."
            
        if messagebox.askyesno("Confirmar Exclusão", confirmation_message):
            try:
                for item_id_selecionado in selecionado:
                    self.db_manager.delete_sale(int(item_id_selecionado)) # Exclui do banco de dados
                
                messagebox.showinfo("Sucesso", f"{len(selecionado)} venda(s) excluída(s) com sucesso!")
                self.atualizar_tabela() # Atualiza a tabela visual
                self.limpar_campos_e_resetar_edicao() # Limpa o formulário e reseta o estado
            except ValueError:
                messagebox.showerror("Erro", "ID de venda inválido.\nSelecione uma linha de venda válida para exclusão.")
            except Exception as e:
                error_message = str(e).replace('%', '%%')
                messagebox.showerror("Erro", f"Erro ao excluir venda(s): {error_message}\nOcorreu um problema ao tentar excluir a(s) venda(s) selecionada(s).")

    def exportar_excel(self):
        """
        Exporta todos os dados de vendas para um arquivo Excel na aba 'Vendas'.
        Se o arquivo já existir, ele preserva a aba 'Encomendas'.
        Inclui uma linha de total e formata as colunas de valor.
        """
        if not messagebox.askyesno("Confirmar Exportação", "Deseja realmente exportar todas as vendas para um arquivo Excel?\nIsso pode levar alguns segundos se houver muitos dados."):
            return

        excel_file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")

        # Carrega os dados existentes da aba 'Encomendas' se o arquivo existir
        df_encomendas_existente = pd.DataFrame()
        if os.path.exists(excel_file_path):
            try:
                df_encomendas_existente = pd.read_excel(excel_file_path, sheet_name='Encomendas')
            except ValueError:
                # A aba 'Encomendas' pode não existir, o que é normal
                pass
            except Exception as e:
                messagebox.showwarning("Aviso de Leitura de Excel", f"Não foi possível ler a aba 'Encomendas' do arquivo Excel: {e}\nContinuando com a exportação de vendas.")

        try:
            # Seleciona 'quantidade' também da tabela 'vendas'
            df_vendas = pd.read_sql_query("SELECT id, data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas", self.db_manager.conn)
            if df_vendas.empty:
                messagebox.showinfo("Informação", "Não há vendas para exportar.")
                # Se não há vendas, mas há encomendas, ainda podemos reescrever o arquivo com apenas encomendas
                if df_encomendas_existente.empty:
                    return # Se não há nem vendas nem encomendas, não há nada para exportar

            total_vendas = df_vendas["preco_final"].sum()
            
            # Renomeia as colunas do DataFrame para nomes mais amigáveis no Excel
            df_vendas.rename(columns={
                'id': 'ID da Venda',
                'data_hora': 'Data e Hora',
                'nome_cliente': 'Nome do Cliente',
                'nome_produto': 'Nome do Produto',
                'quantidade': 'Quantidade', 
                # ALTERADO: 'Valor (R$)' para 'Valor por unidade (R$)'
                'preco': 'Valor por unidade (R$)', 
                'tipo_pagamento': 'Tipo de Pagamento',
                'nome_vendedor': 'Nome do Vendedor',
                'preco_final': 'Preço Final (R$)'
            }, inplace=True)

            # Cria uma linha de total para o DataFrame de vendas
            total_row_df_vendas = pd.DataFrame([{
                "ID da Venda": "", 
                "Data e Hora": "",
                "Nome do Cliente": "",
                "Nome do Produto": "",
                "Quantidade": float('nan'), # Usar NaN para colunas numéricas sem valor
                "Valor por unidade (R$)": float('nan'), # ALTERADO: Nome da coluna no Excel
                "Tipo de Pagamento": "TOTAL GERAL:",
                "Preço Final (R$)": total_vendas,
                "Nome do Vendedor": ""
            }])
            
            df_vendas_final = pd.concat([df_vendas, total_row_df_vendas], ignore_index=True)

            os.makedirs(self.excel_export_folder, exist_ok=True) # Garante que a pasta de exportação exista
            
            with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                # Escreve a aba de Vendas
                df_vendas_final.to_excel(writer, sheet_name='Vendas', index=False)

                # Escreve a aba de Encomendas (existente ou vazia)
                df_encomendas_existente.to_excel(writer, sheet_name='Encomendas', index=False)

                workbook = writer.book
                
                # Formatação para a aba 'Vendas'
                worksheet_vendas = writer.sheets['Vendas']
                money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
                # Aplica formatação de moeda às colunas de valor na aba 'Vendas'
                worksheet_vendas.set_column('F:F', None, money_format) # ALTERADO: Coluna 'Valor por unidade (R$)'
                worksheet_vendas.set_column('H:H', None, money_format) # Coluna 'Preço Final (R$)'

                # Formatação para a aba 'Encomendas' (se houver dados)
                worksheet_encomendas = writer.sheets['Encomendas']
                # Aplica formatação de moeda às colunas de valor na aba 'Encomendas'
                if not df_encomendas_existente.empty:
                    worksheet_encomendas.set_column('E:E', None, money_format) # Coluna 'Valor Unitário (R$)'
                    worksheet_encomendas.set_column('G:G', None, money_format) # Coluna 'Valor Total da Encomenda (R$)'

                # Auto-ajuste das colunas para a aba 'Vendas'
                for i, col in enumerate(df_vendas_final.columns):
                    max_len = max(len(str(col)), df_vendas_final[col].astype(str).map(len).max())
                    worksheet_vendas.set_column(i, i, max_len + 2) # +2 para um pequeno padding

                # Auto-ajuste das colunas para a aba 'Encomendas'
                for i, col in enumerate(df_encomendas_existente.columns):
                    max_len = max(len(str(col)), df_encomendas_existente[col].astype(str).map(len).max())
                    worksheet_encomendas.set_column(i, i, max_len + 2) # +2 para um pequeno padding


            messagebox.showinfo("Exportado", f"Arquivo Excel salvo com sucesso em:\n{excel_file_path}\nTotal das vendas incluído e colunas formatadas!")
        except Exception as e:
            error_message = str(e).replace('%', '%%')
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {error_message}\nVerifique se o arquivo não está aberto em outro programa ou se há espaço em disco suficiente. Certifique-se de que a biblioteca 'openpyxl' esteja instalada (`pip install openpyxl`).")

    def limpar_tabela_visual(self):
        """Limpa todas as linhas da Treeview (apenas visualmente, não do banco de dados)."""
        for row in self.tree.get_children():
            self.tree.delete(row)

    def change_theme(self, event=None):
        """Altera o tema da aplicação com base na seleção do combobox e salva a preferência."""
        selected_theme = self.theme_combobox.get()
        self.style.theme_use(selected_theme)
        self.current_theme_name = selected_theme
        self._save_theme_setting(selected_theme)
        # Ao mudar o tema, atualiza a tabela para aplicar o novo estilo à linha total
        self.atualizar_tabela() 
        # Se o caderno estiver aberto, tenta atualizar o estilo da linha total lá também
        if hasattr(self, 'caderno_window') and self.caderno_window.winfo_exists():
            # Acessa a instância do CadernoVirtual e chama seu método para recalcular/reestilizar o total
            self.caderno_window._calculate_total()


    def _load_theme_setting(self):
        """Carrega o tema salvo do arquivo de configurações. Retorna 'cosmo' se não houver ou houver erro."""
        if not os.path.exists(self.theme_settings_path):
            try:
                with open(self.theme_settings_path, 'w', encoding='utf-8') as f:
                    f.write("cosmo") # Escreve o tema padrão se o arquivo não existir
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o arquivo de configurações do tema: {e}")
            return "cosmo"

        try:
            with open(self.theme_settings_path, 'r', encoding='utf-8') as f:
                theme = f.read().strip()
                if theme in self.style.theme_names(): # Verifica se o tema lido é válido
                    return theme
        except Exception as e:
            messagebox.showerror("Erro ao Carregar Tema", f"Não foi possível carregar o tema salvo: {e}\nO tema padrão será usado.")
        return "cosmo" # Retorna o tema padrão em caso de erro ou tema inválido

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
        self.limpar_tabela_visual() # Limpa a tabela antes de recarregar
        
        search_term = self.search_entry.get().strip()
        vendas = self.db_manager.fetch_all_sales(search_term)
        
        total_geral_preco_final = 0.0

        for venda_db in vendas:
            venda_id = venda_db[0]
            # A ordem dos campos em venda_db é:
            # id, data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor
            
            # Formata os valores para exibição na Treeview
            quantidade_formatada = str(venda_db[4]) if venda_db[4] is not None else "" # Quantidade na posição 4
            preco_formatado = f"{venda_db[5]:.2f}".replace(".",",") if venda_db[5] is not None else "" # Preço na posição 5
            preco_final_formatado = f"{venda_db[7]:.2f}".replace(".",",") if venda_db[7] is not None else "" # Preço final na posição 7

            # A ordem dos valores na tupla deve corresponder à ordem das colunas em self.tree.columns
            # (data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor)
            dados_para_treeview = (
                venda_db[1], # data_hora
                venda_db[2], # nome_cliente
                venda_db[3], # nome_produto
                quantidade_formatada, # quantidade (novo)
                preco_formatado, # preco unitário
                venda_db[6], # tipo_pagamento
                preco_final_formatado, # preco_final
                venda_db[8]  # nome_vendedor
            )
            self.tree.insert('', tk.END, iid=venda_id, values=dados_para_treeview)
            total_geral_preco_final += (venda_db[7] if venda_db[7] is not None else 0.0) # Soma o preco_final

        total_final_formatado = f"{total_geral_preco_final:.2f}".replace(".", ",")
        # Insere a linha de total na Treeview
        self.tree.insert('', tk.END, iid=self.TOTAL_ROW_ID, values=(
            "", # data_hora
            "", # nome_cliente
            "", # nome_produto
            "", # quantidade
            "", # preco unitário
            "TOTAL GERAL:", # tipo_pagamento (usado para rótulo)
            total_final_formatado, # preco_final
            "" # nome_vendedor
        ), tags=('total_row',))
        
        # Configura o estilo para a linha de total
        total_row_bg_color = self.style.lookup('TFrame', 'background')
        self.style.configure('total_row.Treeview', background=total_row_bg_color, font=('Arial', 10, 'bold'))

    def fechar_app(self):
        """Fecha a conexão com o banco de dados antes de fechar a aplicação."""
        self.db_manager.close_connection()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesApp(root)
    root.mainloop()











