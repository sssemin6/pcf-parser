"""OpenAI-compatible HTTPS client for structural reviews."""

from dataclasses import dataclass
import http.client
import json
import os
import re
import ssl
from urllib.parse import urlsplit


_CHAT = "/v1/chat/completions"
_MAX_RESPONSE = 2 * 1024 * 1024
_FENCE = re.compile(r"\s*\x60{3}(?:json)?[ \t]*\r?\n(.*?)\r?\n\x60{3}\s*", re.I | re.S)


class GatewayError(RuntimeError):
    pass


def _env(name: str) -> str:
    return os.environ.get(name, "").strip()


def completion_json(content: str) -> dict:
    if not isinstance(content, str):
        raise GatewayError("LLM response has no message content.")
    fence = _FENCE.fullmatch(content)
    try:
        data = json.loads(fence[1] if fence else content)
    except ValueError:
        raise GatewayError("LLM message is not valid JSON.") from None
    if not isinstance(data, dict):
        raise GatewayError("LLM message must be a JSON object.")
    return data


@dataclass(repr=False)
class GatewayClient:
    host: str
    port: int
    model: str
    key: str
    context: ssl.SSLContext
    timeout: float = 120

    @classmethod
    def from_env(cls) -> "GatewayClient":
        raw = _env("PCF_LLM_URL") or _env("MAILSEARCH_GATEWAY_URL")
        key = _env("PCF_LLM_API_KEY") or _env("MAILSEARCH_GATEWAY_API_KEY")
        model = (_env("PCF_LLM_MODEL") or _env("MAILSEARCH_GATEWAY_REWRITE_MODEL")
                 or next((name.strip() for name in _env("MAILSEARCH_GATEWAY_ANSWER_MODELS").split(",") if name.strip()), ""))
        if not raw or not key or not model:
            raise GatewayError("Set PCF_LLM_URL, PCF_LLM_API_KEY, and PCF_LLM_MODEL or their MAILSEARCH_GATEWAY equivalents.")
        if any(ord(char) <= 32 or ord(char) >= 127 for char in raw):
            raise GatewayError("The gateway URL must be an HTTPS origin.")
        try:
            endpoint = urlsplit(raw)
            port = endpoint.port
        except ValueError:
            raise GatewayError("The gateway URL or port is invalid.") from None
        if (endpoint.scheme != "https" or not endpoint.hostname or "@" in endpoint.netloc
                or "?" in raw or "#" in raw or port == 0
                or endpoint.path not in {"", "/", "/v1", "/v1/", _CHAT, _CHAT + "/"}):
            raise GatewayError("The gateway URL must be an HTTPS origin or /v1/chat/completions endpoint.")
        if "\n" in key or "\r" in key:
            raise GatewayError("The gateway API key contains a line break.")
        verify = _env("PCF_LLM_TLS_VERIFY") or _env("MAILSEARCH_GATEWAY_TLS_VERIFY") or "1"
        ca_file = _env("PCF_LLM_CA_FILE") or _env("MAILSEARCH_GATEWAY_CA_FILE")
        if verify not in {"0", "1"} or (verify == "0" and ca_file):
            raise GatewayError("Set TLS verification to 0 or 1. A CA file requires TLS verification.")
        try:
            context = ssl.create_default_context()
            if ca_file:
                context.load_verify_locations(cafile=ca_file)
            if verify == "0":
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
        except OSError:
            raise GatewayError("Could not load the gateway TLS CA configuration.") from None
        return cls(endpoint.hostname, port or 443, model, key, context)

    def complete(self, instructions: str, prompt: str) -> dict:
        body = json.dumps({
            "model": self.model,
            "messages": [{"role": "system", "content": instructions}, {"role": "user", "content": prompt}],
            "temperature": 0,
            "stream": False,
        }, ensure_ascii=False).encode("utf-8")
        connection = http.client.HTTPSConnection(self.host, self.port, timeout=self.timeout, context=self.context)
        try:
            connection.request("POST", _CHAT, body=body, headers={
                "Authorization": f"Bearer {self.key}", "Content-Type": "application/json", "Accept": "application/json",
            })
            response = connection.getresponse()
            if response.status != 200:
                raise GatewayError(f"LLM gateway returned HTTP {response.status}.")
            payload = response.read(_MAX_RESPONSE + 1)
        except (OSError, http.client.HTTPException, ValueError) as exc:
            raise GatewayError(f"LLM gateway request failed ({type(exc).__name__}).") from None
        finally:
            connection.close()
        if len(payload) > _MAX_RESPONSE:
            raise GatewayError("LLM response exceeded the size limit.")
        try:
            content = json.loads(payload)["choices"][0]["message"]["content"]
        except (ValueError, KeyError, IndexError, TypeError):
            raise GatewayError("LLM gateway response does not contain choices[0].message.content.") from None
        return completion_json(content)
