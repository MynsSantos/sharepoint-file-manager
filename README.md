# SharePoint File Manager

Script Python para gerenciar arquivos e pastas do SharePoint/OneDrive usando Microsoft Graph API.

## Funcionalidades

- Autenticação segura via MSAL (Microsoft Authentication Library)
- Obtenção de informações detalhadas de arquivos/pastas através de links compartilhados
- Listagem de conteúdo de pastas
- Suporte para Drive ID e Item ID
- Tratamento robusto de erros

## Tecnologias Utilizadas

- Python 3.8+
- Microsoft Graph API
- MSAL (Microsoft Authentication Library)
- Requests

## Pré-requisitos

1. Registro de aplicativo no Azure AD
2. Permissões necessárias:
   - `Files.Read.All`
   - `Sites.Read.All`

## Instalação

```bash
# Clone o repositório
git clone https://github.com/MynsSantos/sharepoint-file-manager.git
cd sharepoint-file-manager

# Instale as dependências
pip install -r requirements.txt

# Configure as variáveis de ambiente
cp .env.example .env
# Edite o arquivo .env com suas credenciais
```

## Configuração

Crie um arquivo `.env` na raiz do projeto:

```env
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
TENANT_ID=your-tenant-id
```

## Uso

```bash
python sharepoint_manager.py
```

O script solicitará um link compartilhado do OneDrive/SharePoint e exibirá:
- Nome do arquivo/pasta
- Drive ID
- Item ID
- URL web
- Tipo (pasta ou arquivo)
- Tamanho
- Datas de criação e modificação

Para pastas, oferece a opção de listar todo o conteúdo.

## Exemplo de Uso

```python
from sharepoint_manager import SharePointManager

# Inicializar gerenciador
manager = SharePointManager()
manager.authenticate()

# Obter informações de um link
info = manager.get_drive_info_from_url("https://...")
print(info)

# Listar conteúdo de uma pasta
contents = manager.list_folder_contents(drive_id, item_id)
for item in contents:
    print(item['name'])
```

## Estrutura do Projeto

```
sharepoint-file-manager/
├── sharepoint_manager.py    # Script principal
├── requirements.txt          # Dependências
├── .env.example             # Exemplo de configuração
└── README.md                # Documentação
```

## Casos de Uso

- Automação de backup de arquivos
- Integração com sistemas empresariais
- Auditoria de estruturas de pastas
- Migração de dados entre ambientes

## Autor

Yasmim Santos

## Licença

MIT License
