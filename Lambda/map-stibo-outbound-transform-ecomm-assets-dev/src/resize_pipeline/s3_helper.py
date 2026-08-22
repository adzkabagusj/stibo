"""
AWS S3 helpers for the Ecomm Asset Resizing Pipeline.

All operations use a single bucket: config.S3_BUCKET

  download_bytes(key)        — download any object from the bucket
  upload_zip(filename, data) — upload a ZIP to transformed/ecomm-assets/
"""

import boto3
from botocore.exceptions import ClientError

from . import config

# ── Module-level S3 client (reused across warm invocations) ───────────────────
_s3 = boto3.client("s3")


def download_bytes(key: str) -> bytes:
    """
    Download an object from the pipeline S3 bucket and return its raw bytes.

    Parameters
    ----------
    key : str
        S3 object key within config.S3_BUCKET
        e.g. "step-export/ecomm-assets/1304550.xml"
             "Image Requirements.xlsx"

    Returns
    -------
    bytes
        Raw object content.

    Raises
    ------
    RuntimeError
        If the object does not exist or cannot be read.
    """
    bucket = config.S3_BUCKET

    try:
        response = _s3.get_object(Bucket=bucket, Key=key)
        return response["Body"].read()

    except ClientError as exc:
        error_code = exc.response["Error"]["Code"]

        if error_code in ("NoSuchKey", "NoSuchBucket"):
            raise RuntimeError(
                f"S3 object not found: s3://{bucket}/{key}. "
                "Check that the key is correct and the object exists."
            ) from exc

        raise RuntimeError(
            f"S3 download failed for s3://{bucket}/{key}: {exc}"
        ) from exc


def upload_zip(filename: str, data: bytes) -> str:
    """
    Upload a ZIP archive to the output prefix in the pipeline S3 bucket.

    Output path:  s3://{S3_BUCKET}/transformed/ecomm-assets/{filename}

    Parameters
    ----------
    filename : str
        ZIP filename only — no path prefix.
        e.g. "800x800-2026-06-09_10.30.00.zip"
    data : bytes
        Raw ZIP content.

    Returns
    -------
    str
        The full S3 key of the uploaded object.

    Raises
    ------
    RuntimeError
        If the upload fails for any reason.
    """
    bucket = config.S3_BUCKET
    key    = f"{config.S3_OUTPUT_PREFIX}/{filename}"

    try:
        _s3.put_object(
            Bucket=bucket,
            Key=key,
            Body=data,
            ContentType="application/zip",
        )
        size_kb = len(data) // 1024
        print(f"  ✓ Uploaded s3://{bucket}/{key}  ({size_kb} KB)")
        return key

    except ClientError as exc:
        raise RuntimeError(
            f"S3 upload failed for s3://{bucket}/{key}: {exc}"
        ) from exc