#%%

# -*- coding: utf-8 -*-
# Programa de Gestão de Vendas com SQLite, Pandas e Tkinter (Versão Refatorada e Aprimorada)
# Alterações:
# - Caderno de Encomendas e Anotações agora usam o banco de dados SQLite em vez de arquivos .txt,
#   aumentando a robustez e a segurança dos dados.
# - Lógica de exportação para Excel foi centralizada para evitar duplicação de código.
# - Melhorias na organização do código e na experiência do usuário.

import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import os
from ttkbootstrap import Style
from datetime import datetime
import subprocess
import ctypes  # Importa para manipular atributos de arquivo no Windows

# --- Classe para Gerenciamento do Banco de Dados ---
class DatabaseManager:
    """
    Gerencia a conexão e as operações com o banco de dados SQLite.
    Responsável por inicializar o DB e todas as operações CRUD para
    vendas, encomendas e anotações.
    """
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._initialize_db()

    def _initialize_db(self):
        """
        Inicializa a conexão com o banco de dados e cria as tabelas se não existirem.
        Oculta o arquivo do banco de dados no Windows.
        """
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            
            # 1. Cria a tabela 'vendas'
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

            # 2. Cria a tabela 'encomendas'
            self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS encomendas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_hora_registro TEXT,
                nome_cliente TEXT,
                produto TEXT,
                quantidade INTEGER,
                valor_unitario REAL,
                data_entrega TEXT
            )
            """)

            # 3. Cria a tabela 'anotacoes'
            self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS anotacoes (
                id INTEGER PRIMARY KEY DEFAULT 1, -- Sempre usará o ID 1 para ter um único registro
                conteudo TEXT
            )
            """)
            # Garante que a linha de anotações exista
            self.cursor.execute("INSERT OR IGNORE INTO anotacoes (id, conteudo) VALUES (1, '')")

            self.conn.commit()

            # Ocultar o arquivo do banco de dados (específico para Windows)
            if os.name == 'nt' and os.path.exists(self.db_path):
                try:
                    ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x02) # Atributo HIDDEN
                except Exception as e:
                    print(f"Não foi possível ocultar o arquivo do banco de dados: {e}")

        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao inicializar: {e}")
            exit()

    # --- Métodos para Vendas ---
    def insert_sale(self, data):
        """Insere uma nova venda no banco de dados."""
        try:
            self.cursor.execute("""
            INSERT INTO vendas (nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor, data_hora)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (data["nome_cliente"], data["nome_produto"], data["quantidade"], data["preco"],
                  data["tipo_pagamento"], data["preco_final"], data["nome_vendedor"], data["data_hora"]))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao inserir venda: {e}")

    def update_sale(self, sale_id, data):
        """Atualiza uma venda existente."""
        try:
            self.cursor.execute("""
            UPDATE vendas
            SET nome_cliente=?, nome_produto=?, quantidade=?, preco=?, tipo_pagamento=?, preco_final=?, nome_vendedor=?, data_hora=?
            WHERE id=?
            """, (data["nome_cliente"], data["nome_produto"], data["quantidade"], data["preco"],
                  data["tipo_pagamento"], data["preco_final"], data["nome_vendedor"], data["data_hora"], sale_id))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao atualizar venda: {e}")

    def delete_sale(self, sale_id):
        """Exclui uma venda pelo ID."""
        try:
            self.cursor.execute("DELETE FROM vendas WHERE id=?", (sale_id,))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao excluir venda: {e}")

    def fetch_all_sales(self, search_term=""):
        """Busca todas as vendas, opcionalmente filtrando por um termo de busca."""
        try:
            query = "SELECT id, data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas"
            params = []
            if search_term:
                search_pattern = f"%{search_term}%"
                query += " WHERE nome_cliente LIKE ? OR nome_produto LIKE ? OR nome_vendedor LIKE ?"
                params.extend([search_pattern, search_pattern, search_pattern])
            query += " ORDER BY id DESC"
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar vendas: {e}")
            return []

    def fetch_sale_by_id(self, sale_id):
        """Busca uma venda específica pelo ID."""
        try:
            self.cursor.execute("SELECT id, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, nome_vendedor FROM vendas WHERE id=?", (sale_id,))
            return self.cursor.fetchone()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar venda por ID: {e}")
            return None

    # --- Métodos para Encomendas ---
    def insert_encomenda(self, data):
        """Insere uma nova encomenda."""
        try:
            self.cursor.execute("""
            INSERT INTO encomendas (data_hora_registro, nome_cliente, produto, quantidade, valor_unitario, data_entrega)
            VALUES (?, ?, ?, ?, ?, ?)
            """, (data["data_hora_registro"], data["nome_cliente"], data["produto"],
                  data["quantidade"], data["valor_unitario"], data["data_entrega"]))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao inserir encomenda: {e}")

    def update_encomenda(self, encomenda_id, data):
        """Atualiza uma encomenda."""
        try:
            self.cursor.execute("""
            UPDATE encomendas SET nome_cliente=?, produto=?, quantidade=?, valor_unitario=?, data_entrega=?
            WHERE id=?
            """, (data["nome_cliente"], data["produto"], data["quantidade"],
                  data["valor_unitario"], data["data_entrega"], encomenda_id))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao atualizar encomenda: {e}")

    def delete_encomenda(self, encomenda_id):
        """Exclui uma encomenda."""
        try:
            self.cursor.execute("DELETE FROM encomendas WHERE id=?", (encomenda_id,))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao excluir encomenda: {e}")
    
    def clear_all_encomendas(self):
        """Limpa todas as encomendas da tabela."""
        try:
            self.cursor.execute("DELETE FROM encomendas")
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao limpar encomendas: {e}")

    def fetch_all_encomendas(self):
        """Busca todas as encomendas."""
        try:
            self.cursor.execute("SELECT id, data_hora_registro, nome_cliente, produto, quantidade, valor_unitario, data_entrega FROM encomendas ORDER BY id DESC")
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar encomendas: {e}")
            return []

    # --- Métodos para Anotações ---
    def fetch_anotacoes(self):
        """Busca o conteúdo das anotações."""
        try:
            self.cursor.execute("SELECT conteudo FROM anotacoes WHERE id=1")
            result = self.cursor.fetchone()
            return result[0] if result else ""
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar anotações: {e}")
            return ""

    def save_anotacoes(self, content):
        """Salva o conteúdo das anotações."""
        try:
            # UPDATE OR INSERT para garantir que o registro exista
            self.cursor.execute("UPDATE anotacoes SET conteudo=? WHERE id=1", (content,))
            self.conn.commit()
        except sqlite3.Error as e:
            self.conn.rollback()
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao salvar anotações: {e}")

    def close_connection(self):
        """Fecha a conexão com o banco de dados e reexibe o arquivo no Windows."""
        if self.conn:
            self.conn.close()
        if os.name == 'nt' and os.path.exists(self.db_path):
            try:
                # Atributo NORMAL
                ctypes.windll.kernel32.SetFileAttributesW(self.db_path, 0x80)
            except Exception as e:
                print(f"Não foi possível reexibir o arquivo do banco de dados: {e}")


# --- Classe para o Caderno Virtual de Encomendas ---
class CadernoVirtual:
    """
    Gerencia a janela e a lógica para as encomendas,
    agora utilizando o banco de dados.
    """
    TOTAL_ROW_ID = "caderno_total_row"

    def __init__(self, parent_root, theme_style, db_manager):
        self.db_manager = db_manager
        self.caderno_window = tk.Toplevel(parent_root)
        self.caderno_window.title("Caderno de Encomendas")
        self.caderno_window.geometry("1100x750")
        self.caderno_window.transient(parent_root)
        self.caderno_window.grab_set()

        self.style = theme_style
        self._setup_styles()

        self._create_widgets()
        self._load_content()
        self._reset_input_fields()
    
    def _setup_styles(self):
        self.style.configure('Caderno.TFrame', background=self.style.lookup('TFrame', 'background'))
        self.style.configure('Caderno.TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Caderno.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Caderno.Treeview.Heading', font=('Arial', 10, 'bold'))
        self.style.configure('Caderno.Treeview', rowheight=25)
        self.style.map('Caderno.Treeview', background=[('selected', self.style.colors.primary)])

    def _create_widgets(self):
        frame_caderno = ttk.Frame(self.caderno_window, padding=10, style='Caderno.TFrame')
        frame_caderno.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame_caderno, text="Suas Encomendas", font=('Arial', 16, 'bold'), style='Caderno.TLabel').pack(pady=(0, 10))

        frame_input = ttk.LabelFrame(frame_caderno, text="Adicionar/Editar Encomenda", padding=10)
        frame_input.pack(fill=tk.X, pady=(0, 10))

        # Campos de entrada
        labels = ["Nome:", "Produto:", "Quantidade:", "Valor por unidade (R$):", "Data de Entrega:"]
        self.entries = {}
        for i, text in enumerate(labels):
            ttk.Label(frame_input, text=text).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(frame_input, width=30)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            self.entries[text] = entry
        
        frame_input.grid_columnconfigure(1, weight=1)

        # Botões de ação
        frame_botoes_input = ttk.Frame(frame_input)
        frame_botoes_input.grid(row=len(labels), column=0, columnspan=2, pady=10, sticky="ew")

        self.btn_add_encomenda = ttk.Button(frame_botoes_input, text="Adicionar Encomenda", command=self._add_encomenda, style='success.TButton')
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

        self.btn_update_encomenda = ttk.Button(frame_botoes_input, text="Atualizar Encomenda", command=self._update_encomenda, style='primary.TButton')
        self.btn_cancel_edit = ttk.Button(frame_botoes_input, text="Cancelar Edição", command=self._reset_input_fields, style='secondary.TButton')

        # Treeview para exibir as encomendas
        cols = ("ID", "Data e Hora", "Nome", "Produto", "Quantidade", "Valor por unidade", "Data Entrega", "Valor Total")
        self.tree = ttk.Treeview(frame_caderno, columns=cols, show="headings", style='Caderno.Treeview')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Configuração das colunas
        headings = {"ID": 50, "Data e Hora": 150, "Nome": 150, "Produto": 180, "Quantidade": 80,
                    "Valor por unidade": 120, "Data Entrega": 120, "Valor Total": 110}
        
        for col, width in headings.items():
            self.tree.heading(col, text=f"{col} (R$)" if "Valor" in col else col)
            self.tree.column(col, width=width, anchor="center", stretch=tk.NO if col in ["ID", "Quantidade"] else tk.YES)

        # Scrollbars
        scrollbar_y = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.selected_item_iid = None

        # Botões de gestão
        frame_botoes_gestao = ttk.Frame(frame_caderno, style='Caderno.TFrame')
        frame_botoes_gestao.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(frame_botoes_gestao, text="Excluir Selecionada(s)", command=self._delete_encomenda, style='danger.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes_gestao, text="Limpar Caderno Completo", command=self._clear_all, style='warning.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes_gestao, text="Fechar Caderno", command=self.caderno_window.destroy, style='secondary.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

    def _load_content(self):
        """Carrega encomendas do banco de dados para a Treeview."""
        self.tree.delete(*self.tree.get_children())
        encomendas = self.db_manager.fetch_all_encomendas()
        for enc in encomendas:
            encomenda_id, data_reg, nome, prod, qtd, val_unit, data_ent = enc
            try:
                valor_total = int(qtd) * float(val_unit)
                self.tree.insert('', tk.END, iid=encomenda_id, values=(
                    encomenda_id, data_reg, nome, prod, qtd,
                    f"{val_unit:.2f}".replace(".", ","),
                    data_ent,
                    f"{valor_total:.2f}".replace(".", ",")
                ))
            except (ValueError, TypeError):
                continue
        self._calculate_total()
    
    def _validate_inputs(self):
        """Valida e retorna os dados dos campos de entrada."""
        nome = self.entries["Nome:"].get().strip()
        produto = self.entries["Produto:"].get().strip()
        quantidade_str = self.entries["Quantidade:"].get().strip()
        valor_unitario_str = self.entries["Valor por unidade (R$):"].get().strip()
        data_entrega = self.entries["Data de Entrega:"].get().strip()

        if not all([nome, produto, quantidade_str, valor_unitario_str, data_entrega]):
            messagebox.showwarning("Atenção", "Todos os campos são obrigatórios.")
            return None

        try:
            quantidade = int(quantidade_str)
            if quantidade <= 0: raise ValueError
        except ValueError:
            messagebox.showwarning("Atenção", "Quantidade deve ser um número inteiro positivo.")
            return None

        try:
            valor_unitario = float(valor_unitario_str.replace(",", "."))
            if valor_unitario < 0: raise ValueError
        except ValueError:
            messagebox.showwarning("Atenção", "Valor por unidade inválido.")
            return None
            
        return {
            "nome_cliente": nome, "produto": produto, "quantidade": quantidade,
            "valor_unitario": valor_unitario, "data_entrega": data_entrega
        }

    def _add_encomenda(self):
        """Adiciona uma nova encomenda ao banco de dados."""
        data = self._validate_inputs()
        if data:
            data["data_hora_registro"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            self.db_manager.insert_encomenda(data)
            self._load_content()
            self._reset_input_fields()
    
    def _update_encomenda(self):
        """Atualiza a encomenda selecionada."""
        if not self.selected_item_iid:
            messagebox.showwarning("Atenção", "Nenhuma encomenda selecionada.")
            return

        data = self._validate_inputs()
        if data:
            self.db_manager.update_encomenda(self.selected_item_iid, data)
            messagebox.showinfo("Sucesso", "Encomenda atualizada!")
            self._load_content()
            self._reset_input_fields()

    def _on_tree_select(self, event):
        """Carrega os dados da encomenda selecionada nos campos de entrada."""
        selected_items = self.tree.selection()
        if not selected_items or selected_items[0] == self.TOTAL_ROW_ID:
            self._reset_input_fields()
            return

        self.selected_item_iid = selected_items[0]
        values = self.tree.item(self.selected_item_iid, 'values')
        
        self.entries["Nome:"].delete(0, tk.END); self.entries["Nome:"].insert(0, values[2])
        self.entries["Produto:"].delete(0, tk.END); self.entries["Produto:"].insert(0, values[3])
        self.entries["Quantidade:"].delete(0, tk.END); self.entries["Quantidade:"].insert(0, values[4])
        self.entries["Valor por unidade (R$):"].delete(0, tk.END); self.entries["Valor por unidade (R$):"].insert(0, values[5])
        self.entries["Data de Entrega:"].delete(0, tk.END); self.entries["Data de Entrega:"].insert(0, values[6])

        self.btn_add_encomenda.pack_forget()
        self.btn_update_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        self.btn_cancel_edit.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)

    def _delete_encomenda(self):
        """Exclui a(s) encomenda(s) selecionada(s)."""
        selected_items = [item for item in self.tree.selection() if item != self.TOTAL_ROW_ID]
        if not selected_items:
            messagebox.showwarning("Atenção", "Selecione uma ou mais encomendas para excluir.")
            return

        msg = f"Tem certeza que deseja excluir {len(selected_items)} encomenda(s)?"
        if messagebox.askyesno("Confirmar Exclusão", msg):
            for item_id in selected_items:
                self.db_manager.delete_encomenda(item_id)
            messagebox.showinfo("Sucesso", f"{len(selected_items)} encomenda(s) excluída(s).")
            self._load_content()
            self._reset_input_fields()

    def _reset_input_fields(self):
        """Limpa campos de entrada e redefine botões."""
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        
        self.selected_item_iid = None
        self.btn_update_encomenda.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_add_encomenda.pack(side=tk.LEFT, expand=True, padx=5, fill=tk.X)
        if self.tree.selection():
            self.tree.selection_remove(self.tree.selection())

    def _clear_all(self):
        """Limpa todo o conteúdo do caderno."""
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja limpar TODO o caderno?"):
            self.db_manager.clear_all_encomendas()
            messagebox.showinfo("Sucesso", "O caderno foi limpo.")
            self._load_content()
            self._reset_input_fields()

    def _calculate_total(self):
        """Calcula e exibe o total na Treeview."""
        if self.tree.exists(self.TOTAL_ROW_ID):
            self.tree.delete(self.TOTAL_ROW_ID)

        total_valor = sum(float(self.tree.item(item, 'values')[7].replace(",", "."))
                          for item in self.tree.get_children()
                          if item != self.TOTAL_ROW_ID)

        total_formatado = f"{total_valor:.2f}".replace(".", ",")
        self.tree.insert('', tk.END, iid=self.TOTAL_ROW_ID, values=(
            "", "", "", "", "", "", "TOTAL GERAL:", total_formatado
        ), tags=('total_row',))
        
        self.tree.tag_configure('total_row', background=self.style.lookup('TFrame', 'background'), font=('Arial', 10, 'bold'))

# --- Classe para as Anotações Virtuais ---
class AnotacoesVirtual:
    """Janela de anotações, utilizando o banco de dados."""
    def __init__(self, parent_root, theme_style, db_manager):
        self.db_manager = db_manager
        self.anotacoes_window = tk.Toplevel(parent_root)
        self.anotacoes_window.title("Anotações")
        self.anotacoes_window.geometry("500x600")
        self.anotacoes_window.transient(parent_root)
        self.anotacoes_window.grab_set()
        self.anotacoes_window.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.style = theme_style
        self._create_widgets()
        self._load_content()
    
    def _create_widgets(self):
        frame_anotacoes = ttk.Frame(self.anotacoes_window, padding=10)
        frame_anotacoes.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame_anotacoes, text="Suas Anotações", font=('Arial', 18, 'bold')).pack(pady=(0, 10))

        self.text_area = tk.Text(frame_anotacoes, wrap=tk.WORD, font=('Arial', 12), relief=tk.SOLID, bd=1)
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        frame_botoes = ttk.Frame(frame_anotacoes)
        frame_botoes.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(frame_botoes, text="Salvar Anotações", command=self._save_content, style='info.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Limpar", command=self._clear_content, style='danger.TButton').pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(frame_botoes, text="Fechar", command=self._on_closing, style='secondary.TButton').pack(side=tk.RIGHT, expand=True, padx=5)

    def _load_content(self):
        """Carrega anotações do banco de dados."""
        content = self.db_manager.fetch_anotacoes()
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, content)

    def _save_content(self):
        """Salva o conteúdo no banco de dados."""
        content = self.text_area.get(1.0, tk.END).strip()
        self.db_manager.save_anotacoes(content)
        messagebox.showinfo("Salvo", "Anotações salvas com sucesso!", parent=self.anotacoes_window)

    def _clear_content(self):
        """Limpa o campo de texto e salva."""
        if messagebox.askyesno("Confirmar", "Limpar todas as anotações?", parent=self.anotacoes_window):
            self.text_area.delete(1.0, tk.END)
            self._save_content()

    def _on_closing(self):
        """Salva ao fechar e destrói a janela."""
        self._save_content()
        self.anotacoes_window.destroy()

# --- Classe Principal da Aplicação ---
class SalesApp:
    """Classe principal da aplicação de gestão de vendas."""
    TOTAL_ROW_ID = "total_row_id"

    def __init__(self, root):
        self.root = root
        self._setup_paths_and_dirs()
        self._setup_style()
        
        self.root.title("Sistema de Cadastro de Vendas")
        self.root.geometry("1050x800")

        self.db_manager = DatabaseManager(self.db_path)
        self.id_venda_em_edicao = None

        self._create_widgets()
        self.atualizar_tabela()
        self.limpar_campos_e_resetar_edicao()
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_app)

    def _setup_paths_and_dirs(self):
        """Define os caminhos para os arquivos e diretórios de dados."""
        try:
            # Tenta usar um diretório interno junto ao executável
            app_dir = os.path.dirname(os.path.abspath(__file__))
            data_dir = os.path.join(app_dir, "_internal_data")
            os.makedirs(data_dir, exist_ok=True)
            if not os.access(data_dir, os.W_OK):
                raise OSError("Diretório de dados não é gravável.")
        except OSError:
            # Como alternativa, usa a pasta Documentos do usuário
            data_dir = os.path.join(os.path.expanduser("~"), "Documents", "GestaoVendasData")
            os.makedirs(data_dir, exist_ok=True)
            messagebox.showwarning("Aviso de Caminho", f"Os dados serão salvos em: {data_dir}")

        self.db_path = os.path.join(data_dir, "vendas_gestao.db")
        self.theme_settings_path = os.path.join(data_dir, "theme_setting.txt")
        self.excel_export_folder = os.path.join(os.path.expanduser("~"), "Documents", "Vendas_Padaria_Exportadas")
        os.makedirs(self.excel_export_folder, exist_ok=True)

    def _setup_style(self):
        """Configura o tema ttkbootstrap."""
        self.current_theme_name = self._load_theme_setting()
        self.style = Style(theme=self.current_theme_name)
        self.style.configure("TButton", padding=(10, 5))

    def _create_widgets(self):
        """Cria e organiza todos os widgets da interface gráfica."""
        # Configuração do grid principal
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        ttk.Label(self.root, text="SISTEMA DE GESTÃO DE VENDAS", font=("Arial", 28, "bold")).grid(row=0, column=0, pady=(10, 15))

        # --- Seletor de Tema ---
        frame_theme = ttk.Frame(self.root)
        frame_theme.grid(row=1, column=0, sticky="e", padx=10)
        ttk.Label(frame_theme, text="Tema:").pack(side=tk.LEFT, padx=5)
        self.available_themes = sorted(list(self.style.theme_names()))
        self.theme_combobox = ttk.Combobox(frame_theme, values=self.available_themes, state="readonly", width=15)
        self.theme_combobox.set(self.current_theme_name)
        self.theme_combobox.pack(side=tk.LEFT)
        self.theme_combobox.bind("<<ComboboxSelected>>", self.change_theme)

        # --- Formulário de Venda ---
        frame_formulario = ttk.LabelFrame(self.root, text="Registro de Venda", padding=10)
        frame_formulario.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        frame_formulario.grid_columnconfigure(1, weight=1)
        
        labels_info = ["Nome do Cliente:", "Nome do Produto:", "Quantidade:", "Valor por unidade (R$):", "Tipo de Pagamento:", "Nome do Vendedor:"]
        self.campos = {}
        for i, text in enumerate(labels_info):
            ttk.Label(frame_formulario, text=text).grid(row=i, column=0, sticky="w", padx=5, pady=5)
            if text == "Tipo de Pagamento:":
                widget = ttk.Combobox(frame_formulario, values=["Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Pix"], state="readonly")
            else:
                widget = ttk.Entry(frame_formulario)
            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.campos[text.split(':')[0]] = widget

        frame_botoes_form = ttk.Frame(frame_formulario)
        frame_botoes_form.grid(row=len(labels_info), column=0, columnspan=2, pady=10, sticky="ew")
        self.btn_salvar = ttk.Button(frame_botoes_form, text="Registrar Venda", command=self.salvar_dados, style="success.TButton")
        self.btn_salvar.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.btn_cancelar_edicao = ttk.Button(frame_botoes_form, text="Cancelar Edição", command=self.limpar_campos_e_resetar_edicao, style="secondary.TButton")

        # --- Botões de Gestão ---
        frame_gestao = ttk.LabelFrame(self.root, text="Ações", padding=10)
        frame_gestao.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        
        botoes_gestao = {
            "Carregar para Editar": (self.carregar_para_edicao, "info.TButton"),
            "Excluir Selecionada(s)": (self.excluir_venda_selecionada, "danger.TButton"),
            "Exportar Vendas/Encomendas": (self.exportar_dados, "primary.TButton"),
            "Abrir Planilha": (self.abrir_planilha_excel, "info.TButton"),
            "Calculadora": (lambda: subprocess.Popen('calc.exe') if os.name == 'nt' else subprocess.Popen('gnome-calculator'), "secondary.TButton"),
            "Encomendas": (self.abrir_caderno_encomendas, "success.TButton"),
            "Anotações": (self.abrir_anotacoes, "success.TButton"),
        }
        for i, (text, (command, style)) in enumerate(botoes_gestao.items()):
            ttk.Button(frame_gestao, text=text, command=command, style=style).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
            frame_gestao.grid_columnconfigure(i, weight=1)

        # --- Tabela de Vendas ---
        frame_tabela = ttk.LabelFrame(self.root, text="Vendas Registradas", padding=10)
        frame_tabela.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        frame_tabela.grid_rowconfigure(1, weight=1)
        frame_tabela.grid_columnconfigure(0, weight=1)

        frame_pesquisa = ttk.Frame(frame_tabela, padding=(0, 5))
        frame_pesquisa.grid(row=0, column=0, sticky="ew")
        ttk.Label(frame_pesquisa, text="Pesquisar:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(frame_pesquisa)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.search_entry.bind("<KeyRelease>", lambda e: self.atualizar_tabela())
        ttk.Button(frame_pesquisa, text="Limpar", command=lambda: [self.search_entry.delete(0, tk.END), self.atualizar_tabela()]).pack(side=tk.LEFT, padx=5)

        colunas = ("data_hora", "nome_cliente", "nome_produto", "quantidade", "preco", "tipo_pagamento", "preco_final", "nome_vendedor")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings", selectmode="extended")
        self.tree.grid(row=1, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        headings = {"data_hora": "Data e Hora", "nome_cliente": "Cliente", "nome_produto": "Produto", "quantidade": "Qtd.",
                    "preco": "Valor Unid. (R$)", "tipo_pagamento": "Pagamento", "preco_final": "Preço Final (R$)", "nome_vendedor": "Vendedor"}
        widths = {"data_hora": 140, "quantidade": 60, "preco": 120, "tipo_pagamento": 120, "preco_final": 120}
        
        for col, text in headings.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=widths.get(col, 150), anchor="center")

    def abrir_caderno_encomendas(self):
        CadernoVirtual(self.root, self.style, self.db_manager)

    def abrir_anotacoes(self):
        AnotacoesVirtual(self.root, self.style, self.db_manager)

    def abrir_planilha_excel(self):
        file_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Atenção", "O arquivo Excel ainda não foi exportado.")
            return
        
        try:
            if os.name == 'nt': subprocess.Popen(['start', file_path], shell=True)
            elif os.uname().sysname == 'Darwin': subprocess.Popen(['open', file_path])
            else: subprocess.Popen(['xdg-open', file_path])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a planilha: {e}")

    def salvar_dados(self):
        try:
            dados = {
                "nome_cliente": self.campos["Nome do Cliente"].get().strip(),
                "nome_produto": self.campos["Nome do Produto"].get().strip(),
                "quantidade": self.campos["Quantidade"].get().strip(),
                "preco": self.campos["Valor por unidade (R$)"].get().strip(),
                "tipo_pagamento": self.campos["Tipo de Pagamento"].get().strip(),
                "nome_vendedor": self.campos["Nome do Vendedor"].get().strip(),
            }

            if not all(dados.values()):
                messagebox.showwarning("Atenção", "Todos os campos são obrigatórios.")
                return

            quantidade = int(dados["quantidade"])
            preco_unitario = float(dados["preco"].replace(",", "."))
            if quantidade <= 0 or preco_unitario < 0:
                raise ValueError("Valores devem ser positivos.")

            dados.update({
                "quantidade": quantidade,
                "preco": preco_unitario,
                "preco_final": quantidade * preco_unitario,
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })

            if self.id_venda_em_edicao is not None:
                self.db_manager.update_sale(self.id_venda_em_edicao, dados)
                messagebox.showinfo("Sucesso", "Venda atualizada!")
            else:
                self.db_manager.insert_sale(dados)
                messagebox.showinfo("Sucesso", "Venda registrada!")

            self.limpar_campos_e_resetar_edicao()
            self.atualizar_tabela()

        except ValueError:
            messagebox.showerror("Erro de Validação", "Quantidade ou Preço inválidos. Verifique os valores inseridos.")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")

    def limpar_campos_e_resetar_edicao(self):
        self.id_venda_em_edicao = None
        for widget in self.campos.values():
            if isinstance(widget, ttk.Combobox):
                widget.set("")
            else:
                widget.delete(0, tk.END)
        self.btn_salvar.config(text="Registrar Venda", style="success.TButton")
        self.btn_cancelar_edicao.pack_forget()

    def carregar_para_edicao(self):
        selecionado = self.tree.selection()
        if not selecionado or len(selecionado) > 1 or selecionado[0] == self.TOTAL_ROW_ID:
            messagebox.showwarning("Atenção", "Selecione uma única venda para editar.")
            return

        venda_id = int(selecionado[0])
        venda = self.db_manager.fetch_sale_by_id(venda_id)
        if venda:
            self.limpar_campos_e_resetar_edicao()
            _, nome_cli, nome_prod, qtd, preco, tipo_pag, nome_vend = venda
            
            self.campos["Nome do Cliente"].insert(0, nome_cli)
            self.campos["Nome do Produto"].insert(0, nome_prod)
            self.campos["Quantidade"].insert(0, str(qtd))
            self.campos["Valor por unidade (R$)"].insert(0, str(preco).replace(".", ","))
            self.campos["Tipo de Pagamento"].set(tipo_pag)
            self.campos["Nome do Vendedor"].insert(0, nome_vend)

            self.id_venda_em_edicao = venda_id
            self.btn_salvar.config(text="Atualizar Venda", style="primary.TButton")
            self.btn_cancelar_edicao.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

    def excluir_venda_selecionada(self):
        selecionados = [item for item in self.tree.selection() if item != self.TOTAL_ROW_ID]
        if not selecionados:
            messagebox.showwarning("Atenção", "Selecione uma ou mais vendas para excluir.")
            return

        msg = f"Tem certeza que deseja excluir {len(selecionados)} venda(s)?"
        if messagebox.askyesno("Confirmar Exclusão", msg):
            for item_id in selecionados:
                self.db_manager.delete_sale(int(item_id))
            messagebox.showinfo("Sucesso", f"{len(selecionados)} venda(s) excluída(s).")
            self.atualizar_tabela()
            self.limpar_campos_e_resetar_edicao()

    def exportar_dados(self):
        """Exporta Vendas e Encomendas para um arquivo Excel com duas abas."""
        if not messagebox.askyesno("Confirmar", "Deseja exportar todos os dados de Vendas e Encomendas?"):
            return

        excel_path = os.path.join(self.excel_export_folder, "vendas_padaria_exportadas.xlsx")
        try:
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                # Exportar Vendas
                df_vendas = pd.read_sql_query("SELECT id, data_hora, nome_cliente, nome_produto, quantidade, preco, tipo_pagamento, preco_final, nome_vendedor FROM vendas", self.db_manager.conn)
                if not df_vendas.empty:
                    df_vendas.rename(columns={'id': 'ID Venda', 'preco': 'Valor Unid. (R$)', 'preco_final': 'Preço Final (R$)'}, inplace=True)
                    df_vendas.to_excel(writer, sheet_name='Vendas', index=False)
                    
                # Exportar Encomendas
                df_enc = pd.read_sql_query("SELECT id, data_hora_registro, nome_cliente, produto, quantidade, valor_unitario, data_entrega FROM encomendas", self.db_manager.conn)
                if not df_enc.empty:
                    df_enc['valor_total'] = df_enc['quantidade'] * df_enc['valor_unitario']
                    df_enc.rename(columns={'id': 'ID Encomenda', 'valor_unitario': 'Valor Unid. (R$)', 'valor_total': 'Valor Total (R$)'}, inplace=True)
                    df_enc.to_excel(writer, sheet_name='Encomendas', index=False)
                
                # Formatação
                workbook = writer.book
                money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
                if 'Vendas' in writer.sheets:
                    writer.sheets['Vendas'].set_column('F:F', 12, money_format)
                    writer.sheets['Vendas'].set_column('H:H', 12, money_format)
                if 'Encomendas' in writer.sheets:
                    writer.sheets['Encomendas'].set_column('F:H', 12, money_format)
            
            messagebox.showinfo("Sucesso", f"Dados exportados para:\n{excel_path}")
        except Exception as e:
            messagebox.showerror("Erro de Exportação", f"Não foi possível exportar: {e}")

    def change_theme(self, event=None):
        selected_theme = self.theme_combobox.get()
        self.style.theme_use(selected_theme)
        self.current_theme_name = selected_theme
        self._save_theme_setting(selected_theme)
        self.atualizar_tabela()

    def _load_theme_setting(self):
        try:
            with open(self.theme_settings_path, 'r') as f:
                theme = f.read().strip()
            return theme if theme else "cosmo"
        except FileNotFoundError:
            return "cosmo"

    def _save_theme_setting(self, theme_name):
        with open(self.theme_settings_path, 'w') as f:
            f.write(theme_name)

    def atualizar_tabela(self):
        self.tree.delete(*self.tree.get_children())
        search_term = self.search_entry.get().strip()
        vendas = self.db_manager.fetch_all_sales(search_term)
        
        total_geral = 0.0
        for venda in vendas:
            venda_id, data_h, nome_c, nome_p, qtd, preco_u, tipo_p, preco_f, nome_v = venda
            self.tree.insert('', tk.END, iid=venda_id, values=(
                data_h, nome_c, nome_p, qtd,
                f"{preco_u:.2f}".replace(".", ","),
                tipo_p,
                f"{preco_f:.2f}".replace(".", ","),
                nome_v
            ))
            total_geral += preco_f

        self.tree.insert('', tk.END, iid=self.TOTAL_ROW_ID, values=(
            "", "", "", "", "", "", f"TOTAL: {total_geral:.2f}".replace(".",","), ""
        ), tags=('total_row',))
        self.tree.tag_configure('total_row', background=self.style.lookup('TFrame', 'background'), font=('Arial', 10, 'bold'))

    def fechar_app(self):
        self.db_manager.close_connection()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesApp(root)
    root.mainloop()
