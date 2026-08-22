"""
sftp_by_response_poller.py
============================
AWS Lambda that polls the Blue Yonder (BY) SFTP server (JKTHORMS) for
response files, downloads them, and uploads them to S3 under the
response/blueyonder/ prefix.

Once the file lands in S3 the existing S3-event trigger fires the main
lambda_function.py → handle_by_response() pipeline automatically.

──────────────────────────────────────────────────────────────────────
AUTHENTICATION
──────────────────────────────────────────────────────────────────────
    SFTP credentials are stored in AWS Secrets Manager under the secret
    named by SFTP_SECRET_NAME (default: MAA-STIBO-rms-Secret).

        secret key   → SFTP username
        secret value → SFTP password

──────────────────────────────────────────────────────────────────────
ENVIRONMENT VARIABLES
──────────────────────────────────────────────────────────────────────
    SFTP_SECRET_NAME        Secrets Manager secret name
                            (default: MAA-STIBO-rms-Secret)
    SFTP_HOST               SFTP hostname for JKTHORMS
    SFTP_PORT               SFTP port (default: 22)
    SFTP_REMOTE_DIR         Remote directory to scan on the SFTP server
                            (default: JKTHORMS/Stibo/Inbound/Blueyonder)
    S3_DEST_BUCKET          S3 bucket to upload response files into
    S3_DEST_PREFIX          S3 prefix for uploaded files
                            (default: response/blueyonder/)
    FILE_EXTENSIONS         Comma-separated extensions to pick up
                            (default: .txt,.csv,.json)

NOTE: Files are DELETED from SFTP immediately after a successful S3 upload.
      There is no archive step.
"""

import boto3
import datetime
import json
import logging
import os
import stat
import sys
from pathlib import Path

import paramiko

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("sftp_by_response_poller")

# ─────────────────────────────────────────────────────────────────────────────
# AWS CLIENTS
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")
secretsmanager = boto3.client("secretsmanager")

# ─────────────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ─────────────────────────────────────────────────────────────────────────────
SFTP_SECRET_NAME = os.environ.get("SFTP_SECRET_NAME", "MAA-STIBO-rms-Secret")
SFTP_HOST        = os.environ.get("SFTP_HOST", "10.221.33.21")
SFTP_PORT        = int(os.environ.get("SFTP_PORT", "22"))
SFTP_REMOTE_DIR  = os.environ.get("SFTP_REMOTE_DIR", "/Inbound/Blueyonder_Dev")
S3_DEST_BUCKET   = os.environ.get("OUTPUT_BUCKET", "")
S3_DEST_PREFIX   = os.environ.get("BY_RESPONSE_PREFIX", "response/blueyonder/")
FILE_EXTENSIONS  = tuple(
    ext.strip().lower()
    for ext in os.environ.get("FILE_EXTENSIONS", ".txt,.csv,.json").split(",")
)

TMP = Path("/tmp")

# ─────────────────────────────────────────────────────────────────────────────
# TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────
_TZ_GMT7 = datetime.timezone(datetime.timedelta(hours=7))


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def get_sftp_credentials() -> tuple[str, str]:
    """
    Retrieve SFTP username and password from Secrets Manager.

    The secret stored under MAA-STIBO-rms-Secret is a JSON object where:
        - the key   is the SFTP username
        - the value is the SFTP password
    e.g. {"my_sftp_user": "my_sftp_pass"}
    """
    resp = secretsmanager.get_secret_value(SecretId=SFTP_SECRET_NAME)
    secret = json.loads(resp["SecretString"])

    # The secret is stored as { "<username>": "<password>" }
    if len(secret) == 1:
        username, password = next(iter(secret.items()))
    else:
        # Fallback: standard username/password keys
        username = secret.get("username", "")
        password = secret.get("password", "")

    if not username or not password:
        raise ValueError(
            f"Could not extract SFTP credentials from secret {SFTP_SECRET_NAME}"
        )

    return username, password


def connect_sftp(host: str, port: int, username: str, password: str) -> paramiko.SFTPClient:
    """Open an SFTP connection using password authentication."""
    transport = paramiko.Transport((host, port))
    transport.connect(username=username, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    log.info("Connected to SFTP %s:%d", host, port)
    return sftp


def list_response_files(sftp: paramiko.SFTPClient, remote_dir: str) -> list[str]:
    """List files in the remote directory that match the allowed extensions."""
    try:
        entries = sftp.listdir_attr(remote_dir)
    except FileNotFoundError:
        log.warning("Remote directory not found: %s", remote_dir)
        return []

    files = []
    for entry in entries:
        # Skip directories
        if stat.S_ISDIR(entry.st_mode or 0):
            continue
        if entry.filename.lower().endswith(FILE_EXTENSIONS):
            files.append(entry.filename)

    log.info("Found %d response file(s) in %s", len(files), remote_dir)
    return files


def download_and_upload(
    sftp: paramiko.SFTPClient,
    remote_dir: str,
    filename: str,
    bucket: str,
    prefix: str,
) -> str:
    """Download a file from SFTP to /tmp, then upload it to S3."""
    remote_path = f"{remote_dir.rstrip('/')}/{filename}"
    local_path = TMP / filename

    log.info("[SFTP→S3] Downloading: %s → %s", remote_path, local_path)
    sftp.get(remote_path, str(local_path))
    file_size = local_path.stat().st_size
    log.info("[SFTP→S3] Download complete: %s (%d bytes)", filename, file_size)

    # Add timestamp (GMT+7) to avoid S3 key collisions
    ts     = datetime.datetime.now(_TZ_GMT7).strftime("%Y%m%d_%H%M%S")
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    s3_key = f"{prefix.rstrip('/')}/{stem}_{ts}{suffix}"

    log.info("[SFTP→S3] Uploading to s3://%s/%s", bucket, s3_key)
    s3.upload_file(str(local_path), bucket, s3_key)
    log.info("[SFTP→S3] Upload confirmed: s3://%s/%s (%d bytes)", bucket, s3_key, file_size)

    # Clean up local temp file
    local_path.unlink(missing_ok=True)

    return s3_key


def delete_from_sftp(sftp: paramiko.SFTPClient, remote_dir: str, filename: str):
    """Delete the file from SFTP after it has been successfully uploaded to S3."""
    remote_path = f"{remote_dir.rstrip('/')}/{filename}"
    sftp.remove(remote_path)
    log.info("Deleted from SFTP: %s", remote_path)


# ─────────────────────────────────────────────────────────────────────────────
# LAMBDA HANDLER
# ─────────────────────────────────────────────────────────────────────────────
def poll_sftp(event, context):
    """
    SFTP poll entry point - called by the router in lambda_function.py
    when an EventBridge scheduled event is received.

    1. Connect to JKTHORMS SFTP using credentials from Secrets Manager
       (secret: MAA-STIBO-rms-Secret, key=username, value=password).
    2. List BY response files in SFTP_REMOTE_DIR (default: /outbound).
    3. Download each file and upload to S3 under S3_DEST_PREFIX
       (default: response/blueyonder/).
    4. Move the file to SFTP_ARCHIVE_DIR (default: /outbound/archive),
       or delete it if SFTP_ARCHIVE_DIR is set to empty string.

    The S3 upload fires an S3 event back into this same Lambda,
    which the router directs to by_inbound.handle_by_response() (Step 2).
    """

    if not SFTP_HOST:
        raise ValueError("SFTP_HOST environment variable is required")
    if not S3_DEST_BUCKET:
        raise ValueError("S3_DEST_BUCKET environment variable is required")

    # 1. Get SFTP credentials
    username, password = get_sftp_credentials()

    # 2. Connect to SFTP
    sftp = connect_sftp(SFTP_HOST, SFTP_PORT, username, password)
    uploaded = []

    try:
        # 3. List response files
        files = list_response_files(sftp, SFTP_REMOTE_DIR)

        if not files:
            log.info("No response files to process. Done.")
            return {"statusCode": 200, "files_processed": 0}

        # 4. Download → S3, then delete from SFTP
        for filename in files:
            try:
                s3_key = download_and_upload(
                    sftp, SFTP_REMOTE_DIR, filename, S3_DEST_BUCKET, S3_DEST_PREFIX
                )
                delete_from_sftp(sftp, SFTP_REMOTE_DIR, filename)
                uploaded.append(s3_key)
                log.info(
                    "✓ [BY Response] %s moved from SFTP (%s) → s3://%s/%s",
                    filename, SFTP_REMOTE_DIR, S3_DEST_BUCKET, s3_key,
                )
            except Exception as exc:
                log.error("✗ Failed to process %s: %s", filename, exc)

    finally:
        sftp.close()
        log.info("SFTP connection closed.")

    result = {
        "statusCode": 200,
        "files_processed": len(uploaded),
        "s3_keys": uploaded,
    }
    log.info("Poller complete: %s", json.dumps(result))
    return result