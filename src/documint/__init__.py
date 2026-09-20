"""
DocuMint - Universal Document Automation & Batch Dispatch Engine in Python.
"""

__version__ = "3.0.0"

from documint.core import (
    DocumentConverter,
    UniversalDocumentConverter,
    WinWordConverter,
    EmailSender,
    SMTPSender,
    WinOutlookSender,
    apply_template_replacements,
    replace_placeholders_in_doc,
    validate_placeholders,
    load_dataset,
    is_valid_email,
    process_emails
)
from documint.database import DatabaseManager

__all__ = [
    "__version__",
    "DocumentConverter",
    "UniversalDocumentConverter",
    "WinWordConverter",
    "EmailSender",
    "SMTPSender",
    "WinOutlookSender",
    "apply_template_replacements",
    "replace_placeholders_in_doc",
    "validate_placeholders",
    "load_dataset",
    "is_valid_email",
    "process_emails",
    "DatabaseManager"
]
