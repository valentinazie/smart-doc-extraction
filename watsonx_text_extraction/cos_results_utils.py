"""Shared COS-results helpers used by `download_cos_results.py` and
`delete_cos_results.py`.

Both CLIs list / filter / format the exact same `text_extraction_results/*`
prefix on the watsonx space bucket, so the paginator, matcher and size
formatter live here instead of being duplicated.
"""

from __future__ import annotations

from typing import Iterable, List


COS_PREFIX = "text_extraction_results/"


def list_all_objects(cos_client, bucket: str, prefix: str = COS_PREFIX) -> List[dict]:
    """List every object under `prefix`, transparently paginating past the
    1000-item `list_objects_v2` limit. Returns the raw object dicts from
    `cos_client.list_objects_v2(...)["Contents"]`.
    """
    all_objects: List[dict] = []
    continuation_token = None
    while True:
        kwargs = {"Bucket": bucket, "Prefix": prefix, "MaxKeys": 1000}
        if continuation_token:
            kwargs["ContinuationToken"] = continuation_token
        response = cos_client.list_objects_v2(**kwargs)
        all_objects.extend(response.get("Contents", []))
        if response.get("IsTruncated"):
            continuation_token = response["NextContinuationToken"]
        else:
            break
    return all_objects


def find_matches(
    objects: Iterable[dict], name: str, exact: bool = False
) -> List[dict]:
    """Return the subset of `objects` whose `Key` matches `name`.

    With `exact=True`, the Key must equal `name` literally. Otherwise a
    case-insensitive substring match is used — handy for the `--filter` CLI
    flag where users pass a partial filename like ``"Galaxy"``.
    """
    if exact:
        return [o for o in objects if o["Key"] == name]
    needle = name.lower()
    return [o for o in objects if needle in o["Key"].lower()]


def fmt_size(n: int) -> str:
    """Human-readable byte count."""
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n / (1024 * 1024):.1f} MB"
