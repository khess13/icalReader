"""
PST to Salesforce CSV Extractor
================================
Extracts emails and attachments from an Outlook .pst file and exports
them into multiple relational CSV files ready for Salesforce import.

Output Tables (CSV files):
  - emails.csv           → EmailMessage__c (or Task/EmailMessage object)
  - recipients.csv       → EmailRecipient__c (To/CC/BCC lines)
  - attachments.csv      → Attachment or ContentVersion (metadata)
  - attachment_files/    → Raw attachment binaries (optional save)

Requirements:
    pip install libpff-python pandas tqdm

Usage:
    python pst_to_salesforce.py --pst path/to/file.pst --out ./output
    python pst_to_salesforce.py --pst file.pst --out ./output --save-attachments
"""

import argparse
import csv
import hashlib
import logging
import os
import re
import sys
import uuid
from datetime import timezone
from pathlib import Path

import pandas as pd

try:
    import pypff  # libpff-python
except ImportError:
    sys.exit(
        "ERROR: 'libpff-python' is not installed.\n"
        "Install it with:  pip install libpff-python"
    )

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_str(value) -> str:
    """Return a clean UTF-8 string; never raises."""
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def _safe_dt(dt_obj) -> str:
    """Convert a pypff datetime to an ISO-8601 string (UTC)."""
    if dt_obj is None:
        return ""
    try:
        # pypff returns a datetime-like object; normalise to UTC ISO string
        if hasattr(dt_obj, "astimezone"):
            return dt_obj.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        return str(dt_obj)
    except Exception:
        return ""


def _clean_body(text: str) -> str:
    """Strip excessive whitespace from email body text."""
    if not text:
        return ""
    # Collapse runs of blank lines to a single blank line
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sanitise_filename(name: str) -> str:
    """Remove path-traversal characters from an attachment filename."""
    name = os.path.basename(name)
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name) or "attachment"


# ---------------------------------------------------------------------------
# Core extraction
# ---------------------------------------------------------------------------

class PSTExtractor:
    """Walk a PST file and collect emails, recipients and attachments."""

    def __init__(self, pst_path: str, save_attachments: bool = False, attachment_dir: Path = None):
        self.pst_path = pst_path
        self.save_attachments = save_attachments
        self.attachment_dir = attachment_dir

        # Rows for each CSV table
        self.emails: list[dict] = []
        self.recipients: list[dict] = []
        self.attachments: list[dict] = []

        self._email_count = 0

    # ------------------------------------------------------------------
    def extract(self):
        log.info("Opening PST: %s", self.pst_path)
        pst = pypff.file()
        pst.open(self.pst_path)
        root = pst.get_root_folder()
        self._walk_folder(root, folder_path="")
        pst.close()
        log.info(
            "Extraction complete — %d emails, %d recipients, %d attachments",
            len(self.emails),
            len(self.recipients),
            len(self.attachments),
        )

    # ------------------------------------------------------------------
    def _walk_folder(self, folder, folder_path: str):
        """Recursively walk PST folders."""
        try:
            folder_name = _safe_str(folder.name) or "Root"
        except Exception:
            folder_name = "Unknown"

        current_path = f"{folder_path}/{folder_name}".lstrip("/")

        # Process messages in this folder
        num_messages = folder.number_of_sub_messages
        iterator = range(num_messages)
        if HAS_TQDM and num_messages > 0:
            iterator = tqdm(iterator, desc=f"📂 {current_path[:60]}", unit="msg", leave=False)

        for i in iterator:
            try:
                message = folder.get_sub_message(i)
                self._process_message(message, folder_path=current_path)
            except Exception as exc:
                log.warning("Skipping message %d in '%s': %s", i, current_path, exc)

        # Recurse into sub-folders
        for j in range(folder.number_of_sub_folders):
            try:
                sub_folder = folder.get_sub_folder(j)
                self._walk_folder(sub_folder, folder_path=current_path)
            except Exception as exc:
                log.warning("Skipping sub-folder %d in '%s': %s", j, current_path, exc)

    # ------------------------------------------------------------------
    def _process_message(self, message, folder_path: str):
        email_id = str(uuid.uuid4())
        self._email_count += 1

        # ---- Core fields ------------------------------------------------
        subject      = _safe_str(message.subject)
        sender       = _safe_str(message.sender_name)
        sender_email = _safe_str(message.sender_email_address)
        body_plain   = _clean_body(_safe_str(message.plain_text_body))
        body_html    = _clean_body(_safe_str(message.html_body))
        sent_dt      = _safe_dt(message.delivery_time)
        msg_id       = _safe_str(getattr(message, "message_identifier", ""))
        importance   = _safe_str(getattr(message, "importance", ""))
        has_attach   = message.number_of_attachments > 0

        self.emails.append({
            "Id":                  email_id,          # internal surrogate key
            "MessageId":           msg_id,
            "Subject":             subject,
            "SenderName":          sender,
            "SenderEmail":         sender_email,
            "SentDate":            sent_dt,
            "BodyPlain":           body_plain,
            "BodyHtml":            body_html,
            "Importance":          importance,
            "HasAttachments":      has_attach,
            "AttachmentCount":     message.number_of_attachments,
            "FolderPath":          folder_path,
        })

        # ---- Recipients -------------------------------------------------
        self._extract_recipients(message, email_id)

        # ---- Attachments ------------------------------------------------
        if has_attach:
            self._extract_attachments(message, email_id)

    # ------------------------------------------------------------------
    def _extract_recipients(self, message, email_id: str):
        """Parse To / CC / BCC recipient headers."""
        # pypff exposes recipients via the recipients collection
        try:
            num_recip = message.number_of_recipients
        except Exception:
            num_recip = 0

        for i in range(num_recip):
            try:
                recip = message.get_recipient(i)
                recip_type_raw = getattr(recip, "recipient_type", 0) or 0
                # 0=To, 1=CC, 2=BCC (MAPI values)
                type_map = {0: "To", 1: "CC", 2: "BCC"}
                recip_type = type_map.get(int(recip_type_raw), "To")

                self.recipients.append({
                    "Id":           str(uuid.uuid4()),
                    "EmailId":      email_id,       # FK → emails.Id
                    "RecipientType": recip_type,
                    "DisplayName":  _safe_str(recip.display_name),
                    "EmailAddress": _safe_str(recip.email_address),
                })
            except Exception as exc:
                log.debug("Could not parse recipient %d: %s", i, exc)

    # ------------------------------------------------------------------
    def _extract_attachments(self, message, email_id: str):
        """Extract attachment metadata (and optionally save binaries)."""
        for i in range(message.number_of_attachments):
            try:
                attach = message.get_attachment(i)
                attach_id = str(uuid.uuid4())

                filename   = _safe_str(attach.name) or f"attachment_{i}"
                filename   = _sanitise_filename(filename)
                size_bytes = 0
                sha256     = ""
                saved_path = ""

                data = None
                try:
                    data = attach.read_buffer(attach.size)
                    size_bytes = len(data)
                    sha256 = _sha256(data)
                except Exception:
                    pass

                if self.save_attachments and data and self.attachment_dir:
                    # Save as  <attachment_dir>/<email_id>/<filename>
                    dest_dir = self.attachment_dir / email_id
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest_file = dest_dir / filename
                    # Avoid collisions
                    if dest_file.exists():
                        dest_file = dest_dir / f"{attach_id}_{filename}"
                    dest_file.write_bytes(data)
                    saved_path = str(dest_file)

                mime_type = _safe_str(getattr(attach, "mime_type", ""))
                content_id = _safe_str(getattr(attach, "content_identifier", ""))

                self.attachments.append({
                    "Id":            attach_id,
                    "EmailId":       email_id,      # FK → emails.Id
                    "FileName":      filename,
                    "MimeType":      mime_type,
                    "SizeBytes":     size_bytes,
                    "SHA256":        sha256,
                    "ContentId":     content_id,
                    "SavedFilePath": saved_path,
                })
            except Exception as exc:
                log.warning("Could not extract attachment %d on email %s: %s", i, email_id, exc)


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

# Salesforce field-name mappings  (internal name → Salesforce API name)
# Adjust these to match your actual Salesforce object/field API names.
SF_EMAIL_FIELDS = {
    "Id":              "External_Id__c",
    "MessageId":       "Message_Id__c",
    "Subject":         "Subject__c",
    "SenderName":      "Sender_Name__c",
    "SenderEmail":     "Sender_Email__c",
    "SentDate":        "Sent_Date__c",
    "BodyPlain":       "Body_Plain__c",
    "BodyHtml":        "Body_Html__c",
    "Importance":      "Importance__c",
    "HasAttachments":  "Has_Attachments__c",
    "AttachmentCount": "Attachment_Count__c",
    "FolderPath":      "Folder_Path__c",
}

SF_RECIPIENT_FIELDS = {
    "Id":              "External_Id__c",
    "EmailId":         "Email__r.External_Id__c",   # external ID relationship
    "RecipientType":   "Recipient_Type__c",
    "DisplayName":     "Display_Name__c",
    "EmailAddress":    "Email_Address__c",
}

SF_ATTACHMENT_FIELDS = {
    "Id":              "External_Id__c",
    "EmailId":         "Email__r.External_Id__c",
    "FileName":        "File_Name__c",
    "MimeType":        "Mime_Type__c",
    "SizeBytes":       "Size_Bytes__c",
    "SHA256":          "SHA256__c",
    "ContentId":       "Content_Id__c",
    "SavedFilePath":   "Saved_File_Path__c",
}


def write_csv(rows: list[dict], field_map: dict, out_path: Path):
    """Write rows to a CSV using Salesforce API field names as headers."""
    if not rows:
        log.warning("No rows to write for %s", out_path.name)
        pd.DataFrame(columns=list(field_map.values())).to_csv(out_path, index=False)
        return

    df = pd.DataFrame(rows)

    # Rename columns to Salesforce API names
    df = df.rename(columns=field_map)

    # Keep only mapped columns (drop any unmapped internal columns)
    sf_cols = [c for c in field_map.values() if c in df.columns]
    df = df[sf_cols]

    # Salesforce expects TRUE/FALSE (not Python True/False)
    for col in df.select_dtypes(include="bool").columns:
        df[col] = df[col].map({True: "TRUE", False: "FALSE"})

    df.to_csv(out_path, index=False, quoting=csv.QUOTE_ALL)
    log.info("  ✔ Written %d rows → %s", len(df), out_path)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extract emails from a PST file and export Salesforce-ready CSVs."
    )
    parser.add_argument("--pst",  required=True, help="Path to the .pst file")
    parser.add_argument("--out",  default="./sf_output", help="Output directory (default: ./sf_output)")
    parser.add_argument(
        "--save-attachments",
        action="store_true",
        help="Save raw attachment binaries to disk inside <out>/attachment_files/",
    )
    parser.add_argument(
        "--no-body-html",
        action="store_true",
        help="Omit HTML body from emails CSV (reduces file size)",
    )
    args = parser.parse_args()

    pst_path = Path(args.pst)
    if not pst_path.exists():
        sys.exit(f"ERROR: PST file not found: {pst_path}")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    attach_dir = None
    if args.save_attachments:
        attach_dir = out_dir / "attachment_files"
        attach_dir.mkdir(parents=True, exist_ok=True)

    # ---- Extract --------------------------------------------------------
    extractor = PSTExtractor(
        pst_path=str(pst_path),
        save_attachments=args.save_attachments,
        attachment_dir=attach_dir,
    )
    extractor.extract()

    # ---- Optionally drop HTML body --------------------------------------
    field_map_emails = dict(SF_EMAIL_FIELDS)
    if args.no_body_html:
        for row in extractor.emails:
            row.pop("BodyHtml", None)
        field_map_emails.pop("BodyHtml", None)

    # ---- Write CSVs -----------------------------------------------------
    log.info("Writing CSV files to: %s", out_dir)
    write_csv(extractor.emails,      field_map_emails,     out_dir / "emails.csv")
    write_csv(extractor.recipients,  SF_RECIPIENT_FIELDS,  out_dir / "recipients.csv")
    write_csv(extractor.attachments, SF_ATTACHMENT_FIELDS, out_dir / "attachments.csv")

    # ---- Summary --------------------------------------------------------
    print("\n" + "="*60)
    print("  PST → Salesforce Export Summary")
    print("="*60)
    print(f"  PST file   : {pst_path}")
    print(f"  Output dir : {out_dir.resolve()}")
    print(f"  Emails     : {len(extractor.emails):,}")
    print(f"  Recipients : {len(extractor.recipients):,}")
    print(f"  Attachments: {len(extractor.attachments):,}")
    if args.save_attachments:
        print(f"  Files saved: {attach_dir}")
    print("="*60)
    print("\nSalesforce Import Order:")
    print("  1. emails.csv      → EmailMessage__c  (upsert on External_Id__c)")
    print("  2. recipients.csv  → EmailRecipient__c (upsert on External_Id__c)")
    print("  3. attachments.csv → Attachment__c    (upsert on External_Id__c)")
    print("\nTip: Use Salesforce Data Loader or Bulk API with 'Upsert' mode.")
    print("     Set the 'External ID' field to 'External_Id__c' for all objects.\n")


if __name__ == "__main__":
    main()
