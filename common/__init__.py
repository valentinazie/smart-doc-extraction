"""Shared helpers used by every extraction pipeline under src/.

Keeping the boilerplate (env loading, watsonx credentials, COS client,
LibreOffice conversion) in one place so each pipeline doesn't re-implement
it. Import freely from anywhere under src/:

    from common.config import (
        load_env, get_watsonx_credentials, get_api_client,
        get_space_cos_client, get_master_cos_resource,
    )
    from common.libreoffice import convert_to_pdf, find_libreoffice
"""

from .config import (  # noqa: F401
    REPO_ROOT,
    SRC_ROOT,
    load_env,
    get_watsonx_credentials,
    get_api_client,
    get_space_cos_client,
    get_master_cos_resource,
)
from .libreoffice import convert_to_pdf, find_libreoffice  # noqa: F401
