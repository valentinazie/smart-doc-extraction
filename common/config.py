"""Environment, watsonx.ai credential and COS client helpers.

All of `src/` used to re-implement `load_dotenv()` + `Credentials(...)` +
`APIClient(space_id=...)` + `ibm_boto3.client(...)` in ~6 files. This module
centralizes that so each pipeline just does:

    from common.config import load_env, get_api_client, get_space_cos_client
    load_env()
    client = get_api_client()
    cos, bucket = get_space_cos_client(client)

Two distinct COS endpoints are exposed because the project uses both:

* **Space COS** — the watsonx.ai space-scoped bucket (used for text extraction
  job I/O). Credentials come from `client.spaces.get_details(...)`.

* **Master COS** — a standalone HMAC-authenticated bucket where processed
  images are uploaded for downstream consumption (e.g. signed URLs in agent
  responses). Endpoint, bucket and credentials come from `MASTER_COS_*` env
  vars; nothing is hard-coded.
"""

from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Optional, Tuple

SRC_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = SRC_ROOT.parent


# ---------------------------------------------------------------------------
# .env loading
# ---------------------------------------------------------------------------

def _candidate_env_files() -> list[Path]:
    """Return the list of `.env` paths to try, in priority order.

    Priority: repo root (current), then src/, then the legacy
    `text_extraction/.env` alias (old PPTX pipeline convention).
    """
    return [
        REPO_ROOT / ".env",
        SRC_ROOT / ".env",
        REPO_ROOT / "text_extraction" / ".env",
    ]


def load_env(verbose: bool = False) -> None:
    """Load .env from any of the known locations. Later files override earlier
    ones (dotenv default). Silently no-ops if python-dotenv is missing.
    """
    try:
        from dotenv import load_dotenv
    except ImportError:  # pragma: no cover
        return
    for p in _candidate_env_files():
        if p.exists():
            load_dotenv(p, override=True)
            if verbose:
                print(f"   🔑 loaded env: {p}")


# ---------------------------------------------------------------------------
# watsonx credentials + APIClient
# ---------------------------------------------------------------------------

def get_watsonx_credentials():
    """Return an `ibm_watsonx_ai.Credentials` built from env.

    Expects `WATSONX_URL` and `WATSONX_APIKEY`. Callers should `load_env()`
    first (or have the env injected by the shell).
    """
    from ibm_watsonx_ai import Credentials
    try:
        return Credentials(
            url=os.environ["WATSONX_URL"],
            api_key=os.environ["WATSONX_APIKEY"],
        )
    except KeyError as e:
        raise RuntimeError(
            f"Missing watsonx env var {e.args[0]}. "
            "Run `load_env()` or set it in your shell."
        ) from None


def get_api_client(space_id: Optional[str] = None, *, project_id: Optional[str] = None):
    """Return an `ibm_watsonx_ai.APIClient`.

    If `space_id` is omitted, falls back to env `SPACE_ID` (most pipelines
    operate in a space). Pass `project_id` for project-scoped flows (VLM
    captioning on watsonx.ai).
    """
    from ibm_watsonx_ai import APIClient
    credentials = get_watsonx_credentials()
    sid = space_id if space_id is not None else os.environ.get("SPACE_ID")
    if sid:
        return APIClient(credentials=credentials, space_id=sid)
    pid = project_id if project_id is not None else os.environ.get("WATSONX_PROJECT_ID")
    if pid:
        return APIClient(credentials=credentials, project_id=pid)
    return APIClient(credentials=credentials)


# ---------------------------------------------------------------------------
# COS — space-scoped bucket (watsonx space storage)
# ---------------------------------------------------------------------------

def _default_space_bucket() -> str:
    bucket = os.environ.get("COS_BUCKET_NAME")
    if not bucket:
        raise RuntimeError(
            "COS_BUCKET_NAME is not set. Add it to your .env "
            "(the bucket attached to your watsonx Space)."
        )
    return bucket


def get_space_cos_client(
    api_client=None,
    *,
    bucket: Optional[str] = None,
) -> Tuple[object, str]:
    """Return `(cos_client, bucket_name)` for the watsonx space COS.

    Re-uses `api_client` if provided; otherwise builds one via
    `get_api_client()`. The returned `cos_client` is an `ibm_boto3` S3 client
    authenticated with the "editor" HMAC keys pulled from the space's
    storage properties.
    """
    import ibm_boto3
    if api_client is None:
        api_client = get_api_client()

    space_id = api_client.default_space_id or os.environ.get("SPACE_ID")
    if not space_id:
        raise RuntimeError("No space_id available on api_client or env.")

    cos_creds = api_client.spaces.get_details(space_id=space_id)[
        "entity"]["storage"]["properties"]
    cos_client = ibm_boto3.client(
        service_name="s3",
        endpoint_url=cos_creds["endpoint_url"],
        aws_access_key_id=cos_creds["credentials"]["editor"]["access_key_id"],
        aws_secret_access_key=cos_creds["credentials"]["editor"]["secret_access_key"],
    )
    return cos_client, (bucket or _default_space_bucket())


# ---------------------------------------------------------------------------
# COS — master images bucket (standalone HMAC)
# ---------------------------------------------------------------------------

def get_master_cos_resource(*, as_client: bool = False):
    """Return an `ibm_boto3` resource (default) or client for the master
    images COS bucket (HMAC-authenticated standalone bucket).

    Reads HMAC creds from env:
        MASTER_COS_ENDPOINT, MASTER_COS_ACCESS_KEY, MASTER_COS_SECRET_KEY
    """
    import ibm_boto3
    endpoint = os.environ.get("MASTER_COS_ENDPOINT")
    ak = os.environ.get("MASTER_COS_ACCESS_KEY")
    sk = os.environ.get("MASTER_COS_SECRET_KEY")
    if not all([endpoint, ak, sk]):
        raise RuntimeError(
            "Master COS env vars missing "
            "(MASTER_COS_ENDPOINT / MASTER_COS_ACCESS_KEY / MASTER_COS_SECRET_KEY)."
        )
    factory = ibm_boto3.client if as_client else ibm_boto3.resource
    return factory(
        service_name="s3",
        endpoint_url=endpoint,
        aws_access_key_id=ak,
        aws_secret_access_key=sk,
    )
