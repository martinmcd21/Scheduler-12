import base64
import time
from typing import Dict, Any, Optional
import requests


class GraphConfig:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        scheduler_mailbox: str,
        base_url: str = "https://graph.microsoft.com/v1.0",
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scheduler_mailbox = scheduler_mailbox
        self.base_url = base_url


def with_retry(max_attempts=3, base_delay_s=1.0):
    def decorator(fn):
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(max_attempts):
                try:
                    return fn(*args, **kwargs)
                except Exception as e:
                    last_exc = e
                    time.sleep(base_delay_s * (2 ** attempt))
            raise last_exc
        return wrapper
    return decorator


class GraphClient:
    def __init__(self, cfg: GraphConfig):
        self.cfg = cfg
        self._token: Optional[str] = None

    def _get_token(self) -> str:
        if self._token:
            return self._token

        url = f"https://login.microsoftonline.com/{self.cfg.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": self.cfg.client_id,
            "client_secret": self.cfg.client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=15)
        r.raise_for_status()
        self._token = r.json()["access_token"]
        return self._token

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _request(self, method: str, url: str, params=None, json_body=None):
        r = requests.request(
            method,
            url,
            headers=self._headers(),
            params=params,
            json=json_body,
            timeout=20,
        )
        r.raise_for_status()
        return r.status_code, r.json() if r.content else None

    # ==========================================================
    # 🔴 THIS IS THE IMPORTANT FIX
    # ==========================================================
    @with_retry(max_attempts=3, base_delay_s=1.0)
    def create_event(
        self,
        event_payload: Dict[str, Any],
        send_updates: str = "all",
    ) -> Dict[str, Any]:
        """
        Create a calendar event AND send meeting invites.

        sendUpdates="all" is REQUIRED for Graph to email attendees.
        """
        url = f"{self.cfg.base_url}/users/{self.cfg.scheduler_mailbox}/events"
        params = {"sendUpdates": send_updates}
        _, body = self._request("POST", url, params=params, json_body=event_payload)
        return body or {}

    @with_retry(max_attempts=3, base_delay_s=1.0)
    def patch_event(
        self,
        event_id: str,
        patch_payload: Dict[str, Any],
        send_updates: str = "all",
    ):
        url = f"{self.cfg.base_url}/users/{self.cfg.scheduler_mailbox}/events/{event_id}"
        params = {"sendUpdates": send_updates}
        self._request("PATCH", url, params=params, json_body=patch_payload)

    @with_retry(max_attempts=3, base_delay_s=1.0)
    def send_mail(
        self,
        to,
        subject,
        body,
        attachments=None,
    ):
        url = f"{self.cfg.base_url}/users/{self.cfg.scheduler_mailbox}/sendMail"

        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }

        if attachments:
            message["attachments"] = attachments

        payload = {"message": message, "saveToSentItems": True}
        self._request("POST", url, json_body=payload)
