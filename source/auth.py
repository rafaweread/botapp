"""
Gerencia token OIDC da Extranet O Boticário com cache local e renovação automática.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Tuple

import jwt                   # pip install pyjwt
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PWTimeout,
)

# ---- Configuração ----------------------------------------------------------

_TOKEN_PATH         = Path(".token_cache.json")   # onde salvar o cache
_TOKEN_LEEWAY       = timedelta(seconds=30)       # margem antes do expirar
_URL_LOGIN          = "https://extranet.grupoboticario.com.br/"
_URL_LOGGED_HOME    = "**/home"                  # pode usar coringa do Playwright
_USER_ENV           = "BOTI_USER"
_PASS_ENV           = "BOTI_PASS"

# ---------------------------------------------------------------------------


# =============== Funções internas de cache ================================

def _load_token_from_disk() -> Tuple[str | None, datetime | None]:
    """Tenta ler token + validade do cache. Retorna (token, expires_at) ou (None, None)."""
    if not _TOKEN_PATH.exists():
        return None, None

    try:
        data = json.loads(_TOKEN_PATH.read_text())
        token       = data["value"]
        expires_str = data["expires_at"]
        expires_at  = datetime.fromisoformat(expires_str)

        # válido?
        if expires_at - datetime.now(timezone.utc) > _TOKEN_LEEWAY:
            return token, expires_at
    except (KeyError, ValueError, json.JSONDecodeError):
        pass  # cache corrompido ou campos faltando

    return None, None


def _save_token_to_disk(token: str, expires_at: datetime) -> None:
    data = {"value": token, "expires_at": expires_at.isoformat()}
    _TOKEN_PATH.write_text(json.dumps(data, ensure_ascii=False))


# =============== Login Playwright ================================

def _perform_login() -> Tuple[str, datetime]:
    """
    Faz login interativo com Playwright e devolve (token, expires_at).
    Lê usuário e senha de variáveis de ambiente BOTI_USER / BOTI_PASS.
    """
    usuario = os.environ.get(_USER_ENV, "rafael.lemos")
    senha   = os.environ.get(_PASS_ENV, "@Boti92664669")
    if not usuario or not senha:
        raise RuntimeError(
            f"Credenciais não encontradas em variáveis de ambiente "
            f"{_USER_ENV} / {_PASS_ENV}"
        )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
        "--disable-blink-features=AutomationControlled",  # disfarçar automação
        "--disable-web-security",
        "--allow-running-insecure-content",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ])
        context = browser.new_context(
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/122.0.0.0 Safari/537.36"),
            viewport={"width": 1280, "height": 800},
        )
        page = context.new_page()

        try:
            page.goto(_URL_LOGIN, timeout=60_000)

            page.locator("#signInName").fill(usuario)
            page.locator("#password").fill(senha)
            page.locator("#next").click()
            page.wait_for_url(_URL_LOGGED_HOME, timeout=60_000)

            # Extrai o access_token do localStorage
            token = page.evaluate("""
                () => {
                    const key = Object.keys(localStorage)
                        .find(k => k.startsWith('oidc.user'));
                    if (!key) return null;
                    const data = JSON.parse(localStorage.getItem(key));
                    return data?.access_token || null;
                }
            """)
            if not token:
                raise RuntimeError("Token não encontrado no localStorage")

            # Decodifica sem verificar assinatura – só para pegar o exp
            payload = jwt.decode(token, options={"verify_signature": False})
            expires_at = datetime.fromtimestamp(payload["exp"], tz=timezone.utc)

            return token, expires_at

        except PWTimeout as err:
            raise RuntimeError(f"Timeout durante login: {err}") from err
        finally:
            context.close()
            browser.close()


# =============== Função pública ============================================

def get_token() -> Tuple[str, datetime]:
    """
    Retorna (token, expires_at). Faz login somente se cache estiver ausente ou vencido.
    """
    token, exp = _load_token_from_disk()
    if token:
        return token, exp

    # Cache vazio ou expirado ➜ login
    token, exp = _perform_login()
    _save_token_to_disk(token, exp)
    return token, exp
