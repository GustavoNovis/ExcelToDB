Ferramenta desenvolvida em Python com interface gráfica intuitiva que permite importar dados de arquivos Excel (.xlsx, .xls) para bancos de dados MySQL e SQL Server. Ideal para usuários que precisam migrar dados de planilhas para ambientes de banco de dados de forma rápida e eficiente, sem necessidade de conhecimentos técnicos avançados.

✨ Funcionalidades Principais

  🔌 Conexão com Bancos de Dados
  
    - Suporte nativo para MySQL e SQL Server
    - Configuração simplificada de conexão (host, porta, usuário, senha)    
    - Validação automática de conexão    
    - Portas padrão pré-configuradas (3306 para MySQL, 1433 para SQL Server)

  📁 Gerenciamento de Arquivos
  
    - Suporte a múltiplos arquivos Excel em uma única importação
    - Concatenação automática de dados de diferentes arquivos
    - Visualização prévia dos dados antes da importação
    - Interface drag-and-drop para seleção de arquivos

  🎯 Controle de Importação
  
    - Seletor de colunas para escolher quais dados importar
    
    - Detecção automática de tipos de dados (inteiro, float, texto, data)
    
    - Criação automática de tabelas no banco de dados
    
    - Atualização de tabelas existentes

  💾 Configuração e Persistência
  
    - Salvar e carregar configurações de conexão
    
    - Lembrar preferências do usuário
    
    - Exportar/importar perfis de configuração

  🚀 Como Utilizar
  
    1. Instalação
      bash
      # Instalar dependências necessárias
      pip install pandas mysql-connector-python pyodbc openpyxl
      
    2. Configuração Inicial
      Execute o script Python
  
      Selecione o tipo de banco de dados (MySQL ou SQL Server)
      
      Preencha os dados de conexão:
      
      Host do servidor
      
      Porta (padrão já configurada)
      
      Nome do banco de dados
      
      Usuário e senha
      
      Nome da tabela de destino
    
    3. Importação de Dados
      Clique em "Selecionar Arquivos Excel"
      
      Escolha um ou múltiplos arquivos
      
      Use "Visualizar Dados" para verificar as informações
      
      Selecione as colunas desejadas para importação
      
      Clique em "Importar" para iniciar o processo
      
    4. Gerenciamento de Configurações
      Salve configurações frequentes para reutilização
      
      Carregue configurações salvas anteriormente
      
      Exporte configurações para uso em outros ambientes

🛠️ Tecnologias Utilizadas

    Python 3.12+ - Linguagem de programação principal
    
    Pandas - Manipulação e análise de dados
    
    MySQL Connector - Conexão com bancos MySQL
    
    pyODBC - Conexão com SQL Server
    
    OpenPyXL - Leitura de arquivos Excel
    
    TKinter - Interface gráfica do usuário

  📋 Requisitos do Sistema
  
    Requisitos Mínimos
    
      Python 3.8 ou superior
      
      4GB de RAM
      
      500MB de espaço em disco
      
      Conexão com banco de dados

    Requisitos Recomendados
    
      Python 3.12+
      
      8GB de RAM
      
      1GB de espaço em disco
      
      Conexão estável com banco de dados

  🔧 Configuração de Banco de Dados
  
    MySQL
    
    - Versão 5.7 ou superior
    - ODBC Driver opcional
    - Permissões de CREATE TABLE e INSERT
    
    SQL Server
    
    - Versão 2012 ou superior
    - ODBC Driver 17+ necessário
    - Permissões adequadas para a base de dados

  📊 Formatos Suportados
  
    Arquivos de Entrada
      
        ✅ .xlsx (Excel Open XML Spreadsheet)
        
        ✅ .xls (Excel Binary File Format)

  Bancos de Dados de Saída
  
    ✅ MySQL 5.7+
    
    ✅ MySQL 8.0+
    
    ✅ SQL Server 2012+
    
    ✅ SQL Server 2019+
