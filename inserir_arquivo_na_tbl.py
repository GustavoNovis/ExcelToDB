import pandas as pd
import os
import mysql.connector
import pyodbc
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import *
from datetime import datetime
import json

class DatabaseImporter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Importador de Arquivos Excel para Banco de Dados")
        self.root.geometry("800x600")
        
        # Variáveis de configuração
        self.db_type = tk.StringVar(value="mysql")
        self.host = tk.StringVar()
        self.port = tk.StringVar()
        self.database = tk.StringVar()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.table_name = tk.StringVar()
        self.selected_columns = []
        self.file_paths = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(W, E, N, S))
        
        # Configuração de pesos para redimensionamento
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Tipo de banco de dados
        ttk.Label(main_frame, text="Tipo de Banco:").grid(row=0, column=0, sticky=W, pady=5)
        db_combo = ttk.Combobox(main_frame, textvariable=self.db_type, 
                               values=["mysql", "sqlserver"])
        db_combo.grid(row=0, column=1, sticky=(W, E), pady=5)
        
        # Host
        ttk.Label(main_frame, text="Host:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.host, width=40).grid(row=1, column=1, sticky=(W, E), pady=5)
        
        # Porta
        ttk.Label(main_frame, text="Porta:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.port, width=40).grid(row=2, column=1, sticky=(W, E), pady=5)
        
        # Database
        ttk.Label(main_frame, text="Database:").grid(row=3, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.database, width=40).grid(row=3, column=1, sticky=(W, E), pady=5)
        
        # Usuário
        ttk.Label(main_frame, text="Usuário:").grid(row=4, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.username, width=40).grid(row=4, column=1, sticky=(W, E), pady=5)
        
        # Senha
        ttk.Label(main_frame, text="Senha:").grid(row=5, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.password, show="*", width=40).grid(row=5, column=1, sticky=(W, E), pady=5)
        
        # Nome da tabela
        ttk.Label(main_frame, text="Tabela:").grid(row=6, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.table_name, width=40).grid(row=6, column=1, sticky=(W, E), pady=5)
        
        # Botão para selecionar arquivos
        ttk.Button(main_frame, text="Selecionar Arquivos Excel", 
                  command=self.select_files).grid(row=7, column=0, columnspan=2, pady=10)
        
        # Botão para visualizar dados e selecionar colunas
        ttk.Button(main_frame, text="Visualizar Dados e Selecionar Colunas", 
                  command=self.preview_and_select_columns).grid(row=8, column=0, columnspan=2, pady=10)
        
        # Botão para importar
        ttk.Button(main_frame, text="Importar para o Banco de Dados", 
                  command=self.import_to_database).grid(row=9, column=0, columnspan=2, pady=10)
        
        # Botão para salvar configuração
        ttk.Button(main_frame, text="Salvar Configuração", 
                  command=self.save_config).grid(row=10, column=0, pady=10)
        
        # Botão para carregar configuração
        ttk.Button(main_frame, text="Carregar Configuração", 
                  command=self.load_config).grid(row=10, column=1, pady=10)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para importar")
        self.status_label.grid(row=11, column=0, columnspan=2, pady=10)
        
        # Lista de arquivos selecionados
        ttk.Label(main_frame, text="Arquivos selecionados:").grid(row=12, column=0, sticky=W, pady=5)
        self.file_listbox = tk.Listbox(main_frame, height=5)
        self.file_listbox.grid(row=13, column=0, columnspan=2, sticky=(W, E, N, S), pady=5)
        
        # Configurar portas padrão
        self.set_default_ports()
        
    def set_default_ports(self):
        # Define portas padrão baseadas no tipo de banco selecionado
        if self.db_type.get() == "mysql":
            self.port.set("3306")
        elif self.db_type.get() == "sqlserver":
            self.port.set("1433")
            
        # Atualizar porta quando o tipo de banco mudar
        self.db_type.trace('w', lambda *args: self.set_default_ports())
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if files:
            self.file_paths = list(files)
            self.file_listbox.delete(0, tk.END)
            for file in self.file_paths:
                self.file_listbox.insert(tk.END, os.path.basename(file))
            self.status_label.config(text=f"{len(self.file_paths)} arquivo(s) selecionado(s)")
    
    def preview_and_select_columns(self):
        if not self.file_paths:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo Excel primeiro.")
            return
            
        try:
            # Ler o primeiro arquivo para preview
            df = pd.read_excel(self.file_paths[0])
            columns = df.columns.tolist()
            
            # Janela para seleção de colunas
            select_window = tk.Toplevel(self.root)
            select_window.title("Selecionar Colunas para Importação")
            select_window.geometry("500x400")
            
            ttk.Label(select_window, text="Selecione as colunas para importar:").pack(pady=10)
            
            # Frame para a lista de colunas com scrollbar
            frame = ttk.Frame(select_window)
            frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
            
            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=15)
            scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=listbox.yview)
            listbox.configure(yscrollcommand=scrollbar.set)
            
            for col in columns:
                listbox.insert(tk.END, col)
                
            # Selecionar todas as colunas por padrão
            for i in range(len(columns)):
                listbox.select_set(i)
                
            listbox.pack(side=LEFT, fill=BOTH, expand=True)
            scrollbar.pack(side=RIGHT, fill=Y)
            
            def confirm_selection():
                selected_indices = listbox.curselection()
                self.selected_columns = [columns[i] for i in selected_indices]
                select_window.destroy()
                messagebox.showinfo("Sucesso", f"{len(self.selected_columns)} colunas selecionadas.")
                
            ttk.Button(select_window, text="Confirmar Seleção", command=confirm_selection).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo: {str(e)}")
    
    def get_connection(self):
        try:
            if self.db_type.get() == "mysql":
                conn = mysql.connector.connect(
                    host=self.host.get(),
                    port=int(self.port.get()),
                    database=self.database.get(),
                    user=self.username.get(),
                    password=self.password.get()
                )
                return conn
                
            elif self.db_type.get() == "sqlserver":
                conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.host.get()},{self.port.get()};DATABASE={self.database.get()};UID={self.username.get()};PWD={self.password.get()}'
                conn = pyodbc.connect(conn_str)
                return conn
                
        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Não foi possível conectar ao banco de dados: {str(e)}")
            return None
    
    def import_to_database(self):
        if not self.file_paths:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo Excel.")
            return
            
        if not self.selected_columns:
            messagebox.showerror("Erro", "Selecione as colunas para importar.")
            return
            
        if not all([self.host.get(), self.port.get(), self.database.get(), 
                   self.username.get(), self.password.get(), self.table_name.get()]):
            messagebox.showerror("Erro", "Preencha todas as informações de conexão.")
            return
            
        try:
            # Ler e concatenar todos os arquivos
            dfs = []
            for file_path in self.file_paths:
                df = pd.read_excel(file_path)
                # Selecionar apenas as colunas escolhidas
                df = df[self.selected_columns]
                dfs.append(df)
                
            final_df = pd.concat(dfs, ignore_index=True)
            
            # Conectar ao banco de dados
            conn = self.get_connection()
            if conn is None:
                return
                
            # Inserir dados
            cursor = conn.cursor()
            
            # Criar tabela se não existir
            self.create_table_if_not_exists(cursor, final_df)
            
            # Inserir dados
            self.insert_data(cursor, final_df)
            
            conn.commit()
            conn.close()
            
            self.status_label.config(text=f"Importação concluída! {len(final_df)} registros inseridos.")
            messagebox.showinfo("Sucesso", f"Dados importados com sucesso!\n{len(final_df)} registros inseridos na tabela {self.table_name.get()}.")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a importação: {str(e)}")
    
    def create_table_if_not_exists(self, cursor, df):
        # Gerar comando SQL para criar tabela
        columns_def = []
        for col, dtype in df.dtypes.items():
            if dtype == 'int64':
                sql_type = 'INT'
            elif dtype == 'float64':
                sql_type = 'FLOAT'
            elif dtype == 'datetime64[ns]':
                sql_type = 'DATETIME'
            else:
                sql_type = 'VARCHAR(255)'
            columns_def.append(f"`{col}` {sql_type}")
            
        create_table_sql = f"CREATE TABLE IF NOT EXISTS `{self.table_name.get()}` ({', '.join(columns_def)})"
        cursor.execute(create_table_sql)
    
    def insert_data(self, cursor, df):
        # Gerar placeholders baseados no tipo de banco
        if self.db_type.get() == "mysql":
            placeholders = ", ".join(["%s"] * len(df.columns))
        elif self.db_type.get() == "sqlserver":
            placeholders = ", ".join(["?"] * len(df.columns))
            
        # Gerar SQL de inserção
        columns = ", ".join([f"`{col}`" for col in df.columns])
        insert_sql = f"INSERT INTO `{self.table_name.get()}` ({columns}) VALUES ({placeholders})"
        
        # Converter DataFrame para lista de tuplas
        data_tuples = [tuple(x) for x in df.to_numpy()]
        
        # Inserir dados
        cursor.executemany(insert_sql, data_tuples)
    
    def save_config(self):
        config = {
            'db_type': self.db_type.get(),
            'host': self.host.get(),
            'port': self.port.get(),
            'database': self.database.get(),
            'username': self.username.get(),
            'table_name': self.table_name.get(),
            'selected_columns': self.selected_columns
        }
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(config, f)
            messagebox.showinfo("Sucesso", "Configuração salva com sucesso!")
    
    def load_config(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    config = json.load(f)
                
                self.db_type.set(config.get('db_type', 'mysql'))
                self.host.set(config.get('host', ''))
                self.port.set(config.get('port', ''))
                self.database.set(config.get('database', ''))
                self.username.set(config.get('username', ''))
                self.table_name.set(config.get('table_name', ''))
                self.selected_columns = config.get('selected_columns', [])
                
                # Definir porta padrão se não estiver configurada
                if not self.port.get():
                    self.set_default_ports()
                
                messagebox.showinfo("Sucesso", "Configuração carregada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar configuração: {str(e)}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DatabaseImporter()
    app.run()