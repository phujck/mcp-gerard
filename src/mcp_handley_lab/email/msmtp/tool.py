"""MSMTP email sending provider."""

from pathlib import Path

from pydantic import BaseModel, Field

from mcp_handley_lab.common.process import run_command
from mcp_handley_lab.email.common import mcp


class SendResult(BaseModel):
    """Result of sending an email."""

    status: str = Field(
        default="success",
        description="The status of the send operation, typically 'success'.",
    )
    recipient: str = Field(
        ..., description="The primary recipient's email address (the 'To' field)."
    )
    account_used: str = Field(
        default="", description="The msmtp account used for sending, if specified."
    )
    cc_recipients: list[str] = Field(
        default_factory=list, description="A list of email addresses in the 'Cc' field."
    )
    bcc_recipients: list[str] = Field(
        default_factory=list,
        description="A list of email addresses in the 'Bcc' field.",
    )


def _parse_msmtprc(config_file: str = "") -> list[str]:
    """Parse msmtp config to extract account names."""
    msmtprc_path = Path(config_file) if config_file else Path.home() / ".msmtprc"
    if not msmtprc_path.exists():
        raise FileNotFoundError(f"msmtp configuration not found at {msmtprc_path}")

    accounts = []
    with open(msmtprc_path) as f:
        for line in f:
            line = line.strip()
            if line.startswith("account ") and not line.startswith("account default"):
                account_name = line.split()[1]
                accounts.append(account_name)

    return accounts


def _load_email_file(file_path: str) -> tuple[str, str, str]:
    """Load email file and parse headers (To, Subject) and body."""
    if not Path(file_path).exists():
        raise FileNotFoundError(f"Email file not found: {file_path}")

    with open(file_path, encoding="utf-8") as f:
        content = f.read()

    # Parse email format: headers until empty line, then body
    parts = content.split("\n\n", 1)
    if len(parts) != 2:
        # If no empty line, treat entire content as body
        return "", "", content

    headers_text, body = parts
    to = ""
    subject = ""

    for line in headers_text.split("\n"):
        if line.startswith("To: "):
            to = line[4:].strip()
        elif line.startswith("Subject: "):
            subject = line[9:].strip()

    return to, subject, body


@mcp.tool(
    description="Send email using msmtp. Provide to/subject/body directly or use body_file for pre-composed emails."
)
def send(
    to: str = Field(
        default=None,
        description="The primary recipient's email address. Not needed if using body_file.",
    ),
    subject: str = Field(
        default=None,
        description="The subject line of the email. Not needed if using body_file.",
    ),
    body: str = Field(
        default=None,
        description="The main content (body) of the email. Not needed if using body_file.",
    ),
    body_file: str = Field(
        default=None,
        description="Path to an email file with headers and body. If provided, overrides to/subject/body. Mutually exclusive with body parameter.",
    ),
    account: str = Field(
        default="",
        description="The msmtp account to send from. If empty, the default account is used. Use 'list_accounts' to see options.",
    ),
    cc: list[str] = Field(
        default_factory=list,
        description="List of email addresses for CC recipients.",
    ),
    bcc: list[str] = Field(
        default_factory=list,
        description="List of email addresses for BCC recipients.",
    ),
) -> SendResult:
    """Send an email using msmtp with existing ~/.msmtprc configuration."""
    # Check mutual exclusivity
    if body_file and body:
        raise ValueError("Cannot specify both 'body' and 'body_file'. Choose one.")

    # Load content from body_file if provided
    if body_file:
        file_to, file_subject, file_body = _load_email_file(body_file)
        # Use file values, but allow parameter overrides
        to = to or file_to
        subject = subject or file_subject
        body = file_body

    # Validate that we have required fields
    if not to:
        raise ValueError("Recipient (to) is required.")
    if not subject:
        raise ValueError("Subject is required.")
    if not body:
        raise ValueError("Body is required.")

    email_content = f"To: {to}\n"
    email_content += f"Subject: {subject}\n"

    if cc:
        email_content += f"Cc: {', '.join(cc)}\n"
    if bcc:
        email_content += f"Bcc: {', '.join(bcc)}\n"

    email_content += "\n"
    email_content += body

    cmd = ["msmtp"]
    if account:
        cmd.extend(["-a", account])

    recipients = [to]
    recipients.extend(cc)
    recipients.extend(bcc)

    cmd.extend(recipients)

    input_bytes = email_content.encode()
    stdout, stderr = run_command(cmd, input_data=input_bytes)

    return SendResult(
        recipient=to,
        account_used=account,
        cc_recipients=cc,
        bcc_recipients=bcc,
    )


@mcp.tool(
    description="List available msmtp accounts from ~/.msmtprc configuration. Use to discover valid account names for the send tool."
)
def list_accounts(
    config_file: str = Field(
        default="",
        description="Optional path to the msmtp configuration file. Defaults to `~/.msmtprc`.",
    ),
) -> list[str]:
    """List available msmtp accounts by parsing msmtp config."""
    accounts = _parse_msmtprc(config_file)
    return accounts
