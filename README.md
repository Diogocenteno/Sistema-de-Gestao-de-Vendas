Sistema de Gestão de Vendas - Padaria do Jeff e Pri
🍞 Sobre o Projeto
O Sistema de Gestão de Vendas - Padaria do Jeff e Pri é uma aplicação desktop intuitiva, desenvolvida em Python com tkinter e ttkbootstrap, projetada para otimizar o registro e a organização das vendas diárias de uma padaria ou pequeno comércio. Com ele, você pode facilmente registrar vendas, gerenciar informações de clientes, produtos e vendedores, além de ter recursos adicionais que facilitam o dia a dia.

Este projeto visa proporcionar uma solução prática e eficiente para acompanhar o fluxo de caixa e o histórico de transações, permitindo exportar dados para análise e manter um controle rigoroso de encomendas e anotações importantes.

✨ Funcionalidades Principais
Registro de Vendas Detalhado: Cadastre vendas com informações de cliente, produto, preço, tipo de pagamento e vendedor.
Gestão de Dados:
Edição e Exclusão: Modifique ou remova registros de vendas existentes de forma simples.
Pesquisa Dinâmica: Encontre vendas rapidamente usando a barra de pesquisa integrada, que filtra por cliente, produto ou vendedor.
Exportação para Excel: Exporte todos os dados de vendas para um arquivo .xlsx (Excel), com um resumo do total geral de vendas e colunas formatadas para facilitar a análise.
Calculadora Integrada: Acesse a calculadora do sistema operacional diretamente da aplicação para cálculos rápidos.
Caderno de Encomendas Virtual: Uma seção dedicada para registrar e gerenciar encomendas de clientes, com cálculo automático do valor total das encomendas.
Anotações Virtuais: Um espaço para suas anotações rápidas e importantes, simulando um caderno com linhas e salvamento automático.
Temas Personalizáveis: Escolha entre diversos temas visuais (ttkbootstrap) para personalizar a aparência da aplicação, e a preferência é salva para o próximo uso.
Armazenamento Local e Seguro: Utiliza SQLite para o banco de dados e arquivos de texto para cadernos e anotações, armazenados localmente em uma pasta interna (_internal) ou na pasta Documentos (com fallback automático).
🛠️ Tecnologias Utilizadas
Python 3.x: Linguagem de programação principal.
Tkinter: Biblioteca padrão do Python para criação de interfaces gráficas (GUI).
Ttkbootstrap: Extensão do Tkinter que oferece widgets modernos e temas personalizados, proporcionando uma UI atraente.
SQLite3: Banco de dados leve e embutido para armazenamento de dados das vendas.
Pandas: Biblioteca para manipulação e análise de dados, utilizada na exportação para Excel.
Openpyxl / XlsxWriter: Dependências do Pandas para escrita de arquivos Excel.

🚀 Como Executar o Projeto

![Captura de tela 2025-06-04 075116](https://github.com/user-attachments/assets/9aeb6e62-e872-4f9d-9daf-790829a4189f)

Para rodar este projeto em sua máquina, siga os passos abaixo:

Pré-requisitos
Certifique-se de ter o Python 3.x instalado.

1. Clonar o Repositório

git clone https://github.com/SeuUsuario/Sistema-Gestao-Vendas-Padaria.git
cd Sistema-Gestao-Vendas-Padaria
2. Criar e Ativar um Ambiente Virtual (Recomendado)
Bash

python -m venv venv
# No Windows
.\venv\Scripts\activate
# No macOS/Linux
source venv/bin/activate
3. Instalar as Dependências

pip install ttkbootstrap pandas openpyxl xlsxwriter
4. Executar a Aplicação

python main.py
(Altere main.py para o nome do seu arquivo principal, se for diferente)

📂 Estrutura do Projeto
Sistema-Gestao-Vendas-Padaria/
├── _internal/
│   ├── vendas_padaria.db         # Banco de dados SQLite (oculto no Windows)
│   ├── encomendas_caderno.txt    # Arquivo de texto do caderno de encomendas
│   ├── anotacoes.txt             # Arquivo de texto das anotações
│   └── theme_setting.txt         # Arquivo de configuração do tema
├── main.py                       # Arquivo principal da aplicação
└── README.md                     # Este arquivo
A pasta _internal é criada automaticamente para armazenar os arquivos de dados da aplicação. Se não for possível criá-la ou acessá-la devido a permissões, os arquivos serão salvos na sua pasta Documentos.

🤝 Contribuições
Contribuições são bem-vindas! Se você tiver ideias para melhorias ou encontrar algum bug, sinta-se à vontade para abrir uma issue ou enviar um pull request.

📄 Licença
Este projeto está licenciado sob a Licença MIT - veja o arquivo LICENSE para mais detalhes.
