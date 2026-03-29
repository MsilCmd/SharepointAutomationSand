"""
config/settings.py — Centralised configuration via pydantic-settings.
All values are read from environment variables or a .env file.
"""

from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import Field


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # Azure AD
    azure_tenant_id: str = Field(..., description="Azure AD tenant ID")
    azure_client_id: str = Field(..., description="App registration client ID")
    azure_client_secret: str = Field(..., description="App registration client secret")

    # SharePoint
    sharepoint_site_url: str = Field(..., description="SharePoint site URL")
    sharepoint_username: str | None = Field(None, description="SP admin username (legacy auth)")
    sharepoint_password: str | None = Field(None, description="SP admin password (legacy auth)")

    # Dropbox
    dropbox_app_key: str = Field(..., description="Dropbox app key")
    dropbox_app_secret: str = Field(..., description="Dropbox app secret")
    dropbox_refresh_token: str = Field(..., description="Dropbox refresh token")
    dropbox_access_token: str | None = Field(None, description="Optional static access token")

    # Reporting
    report_output_dir: str = Field("./reports", description="Directory for generated reports")

    @property
    def graph_authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.azure_tenant_id}"

    @property
    def graph_scopes(self) -> list[str]:
        return ["https://graph.microsoft.com/.default"]


# Singleton — import this everywhere
settings = Settings()

