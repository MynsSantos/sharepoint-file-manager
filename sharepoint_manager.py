"""
SharePoint File Manager
=======================
Script para obter informações de arquivos e pastas do SharePoint/OneDrive
usando Microsoft Graph API.

Funcionalidades:
- Autenticação via MSAL (Microsoft Authentication Library)
- Obtenção de Drive ID e Item ID a partir de links compartilhados
- Suporte para pastas e arquivos
- Tratamento de erros e validações

Requisitos:
- Python 3.8+
- msal
- requests
- python-dotenv

Uso:
    python sharepoint_manager.py
"""

import os
import base64
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication


class SharePointManager:
    """Gerenciador de operações com SharePoint via Microsoft Graph API."""

    def __init__(self):
        """Inicializa o gerenciador com credenciais do Azure AD."""
        load_dotenv()
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.tenant_id = os.getenv("TENANT_ID")

        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise ValueError("Credenciais não configuradas no arquivo .env")

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.token = None

    def authenticate(self):
        """Autentica e obtém token de acesso."""
        app = ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret,
        )

        # Tenta obter token silenciosamente
        result = app.acquire_token_silent(self.scope, account=None)
        if not result:
            # Se não conseguir, solicita novo token
            result = app.acquire_token_for_client(scopes=self.scope)

        if "access_token" not in result:
            raise Exception(f"Erro ao obter token: {result.get('error_description')}")

        self.token = result["access_token"]
        return self.token

    def get_drive_info_from_url(self, shared_url: str) -> dict:
        """
        Obtém informações do drive a partir de um link compartilhado.

        Args:
            shared_url: URL compartilhada do OneDrive/SharePoint

        Returns:
            dict: Informações do item (nome, driveId, itemId, tipo, etc)
        """
        if not self.token:
            self.authenticate()

        # Codifica a URL em base64 URL-safe
        encoded = base64.urlsafe_b64encode(shared_url.encode()).decode().rstrip("=")
        share_id = f"u!{encoded}"

        # Consulta a API do Microsoft Graph
        url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
        headers = {"Authorization": f"Bearer {self.token}"}

        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            raise Exception(
                f"Erro ao acessar Graph API: {response.status_code}\n{response.text}"
            )

        data = response.json()

        # Extrai informações relevantes
        return {
            "nome": data.get("name", "Sem nome"),
            "driveId": data["parentReference"]["driveId"],
            "itemId": data["id"],
            "webUrl": data.get("webUrl"),
            "tipo": "pasta" if "folder" in data else "arquivo",
            "tamanho": data.get("size"),
            "criado_em": data.get("createdDateTime"),
            "modificado_em": data.get("lastModifiedDateTime"),
        }

    def list_folder_contents(self, drive_id: str, item_id: str) -> list:
        """
        Lista o conteúdo de uma pasta.

        Args:
            drive_id: ID do drive
            item_id: ID do item (pasta)

        Returns:
            list: Lista de itens dentro da pasta
        """
        if not self.token:
            self.authenticate()

        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
        headers = {"Authorization": f"Bearer {self.token}"}

        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            raise Exception(f"Erro ao listar conteúdo: {response.status_code}")

        return response.json().get("value", [])


def main():
    """Função principal - interface de linha de comando."""
    print("=" * 60)
    print("SharePoint File Manager")
    print("=" * 60)
    print()

    try:
        manager = SharePointManager()
        manager.authenticate()
        print("✓ Autenticação realizada com sucesso!\n")

        link = input("Cole o link compartilhado do OneDrive/SharePoint:\n").strip()

        if not link:
            print("❌ Link não fornecido.")
            return

        print("\nProcessando...")
        resultado = manager.get_drive_info_from_url(link)

        print("\n✓ Dados obtidos com sucesso:")
        print("-" * 60)
        for chave, valor in resultado.items():
            if valor:
                print(f"{chave.capitalize()}: {valor}")

        # Se for uma pasta, perguntar se deseja listar conteúdo
        if resultado["tipo"] == "pasta":
            print("\n" + "-" * 60)
            opcao = input("\nDeseja listar o conteúdo da pasta? (s/n): ").strip().lower()

            if opcao == 's':
                conteudo = manager.list_folder_contents(
                    resultado["driveId"],
                    resultado["itemId"]
                )

                print(f"\nConteúdo da pasta ({len(conteudo)} itens):")
                print("-" * 60)
                for item in conteudo:
                    tipo = "📁" if "folder" in item else "📄"
                    print(f"{tipo} {item['name']}")

    except Exception as e:
        print(f"\n❌ Erro: {e}")
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
