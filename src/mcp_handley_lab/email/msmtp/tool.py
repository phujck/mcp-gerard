"""MSMTP email sending provider with style support."""

from pathlib import Path

from pydantic import BaseModel, Field

from mcp_handley_lab.common.process import run_command
from mcp_handley_lab.email.common import mcp
from mcp_handley_lab.email.style_config import (
    DEFAULT_CONFIG,
    STYLE_CONFIG,
    StyleProfile,
    create_example_body,
    format_guidelines,
)

# Generate dynamic descriptions from loaded style configuration
_first_style = next(iter(STYLE_CONFIG.styles.values()), None)
DEFAULT_STYLE = STYLE_CONFIG.styles.get(STYLE_CONFIG.default_style, _first_style)
GUIDELINES = format_guidelines(DEFAULT_STYLE.guidelines) if DEFAULT_STYLE else ""
EXAMPLE = create_example_body(DEFAULT_STYLE) if DEFAULT_STYLE else ""

# Clean tool description focused on functionality
SEND_DESCRIPTION = "Send email using msmtp. Use draft_email() first to compose styled emails, or provide content directly."


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


@mcp.tool(description=SEND_DESCRIPTION)
def send(
    to: str = Field(
        default=None,
        description="The primary recipient's email address. Not needed if using draft_file.",
    ),
    subject: str = Field(
        default=None,
        description="The subject line of the email. Not needed if using draft_file.",
    ),
    body: str = Field(
        default=None,
        description="The main content (body) of the email. Not needed if using draft_file.",
    ),
    draft_file: str = Field(
        default=None,
        description="Path to a draft email file (created by draft_email with output_file). If provided, overrides to/subject/body.",
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
    # Load content from draft file if provided
    if draft_file:
        draft_content = _load_draft_file(draft_file)
        to = draft_content["to"]
        subject = draft_content["subject"]
        body = draft_content["body"]

    # Validate that we have required fields
    if not to:
        raise ValueError(
            "Recipient (to) is required. Provide either 'to' parameter or 'draft_file'."
        )
    if not subject:
        raise ValueError(
            "Subject is required. Provide either 'subject' parameter or 'draft_file'."
        )
    if not body:
        raise ValueError(
            "Body is required. Provide either 'body' parameter or 'draft_file'."
        )

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


def _load_draft_file(file_path: str) -> dict:
    """Load draft file and parse email headers."""
    if not Path(file_path).exists():
        raise FileNotFoundError(f"Draft file not found: {file_path}")

    with open(file_path, encoding="utf-8") as f:
        content = f.read()

    # Parse email format: headers until empty line, then body
    parts = content.split("\n\n", 1)
    if len(parts) != 2:
        raise ValueError(f"Invalid draft file format: {file_path}")

    headers_text, body = parts
    headers = {}

    for line in headers_text.split("\n"):
        if ":" in line:
            key, value = line.split(":", 1)
            headers[key.strip().lower()] = value.strip()

    return {
        "to": headers.get("to", ""),
        "subject": headers.get("subject", ""),
        "body": body,
    }


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


# Email drafting tool
@mcp.tool(
    description="Draft an email in a specific style. Returns subject and body that can be reviewed and sent via send()."
)
def draft_email(
    style: str = Field(
        default=STYLE_CONFIG.default_style,
        description=f"Email style to use. Available: {list(STYLE_CONFIG.styles.keys())}. Default: {STYLE_CONFIG.default_style}",
    ),
    message_content: str = Field(
        ..., description="The core message content or purpose of the email"
    ),
    recipient: str = Field(..., description="The recipient's name or email address"),
    sender_name: str = Field(
        default=None, description="Sender's name for signature (optional)"
    ),
    output_file: str = Field(
        default=None,
        description="File path to save the draft (optional). If provided, saves as email format for send(draft_file=...)",
    ),
) -> dict:
    """Draft an email in the specified style with subject and body."""
    # Get style profile
    profile = STYLE_CONFIG.styles.get(style)
    if not profile:
        profile = STYLE_CONFIG.styles.get(
            STYLE_CONFIG.default_style, DEFAULT_CONFIG.styles["professional"]
        )
        style = STYLE_CONFIG.default_style

    # Generate subject
    subject = _generate_subject(message_content, profile)

    # Generate body
    body = _generate_body(message_content, recipient, sender_name, profile)

    draft = {
        "subject": subject,
        "body": body,
        "style_used": style,
        "recipient": recipient,
        "sender_name": sender_name,
    }

    # Save to file if requested
    if output_file:
        _save_draft_to_file(draft, output_file)
        draft["draft_file"] = output_file

    return draft


def _generate_subject(message_content: str, profile: StyleProfile) -> str:
    """Generate a subject line following style guidelines."""
    # Simple subject generation - could be enhanced with LLM integration
    subject = f"Re: {message_content}" if message_content else "Email"

    # Apply max length if specified
    if profile.max_subject_len and len(subject) > profile.max_subject_len:
        subject = subject[: profile.max_subject_len - 3] + "..."

    return subject


def _generate_body(
    message_content: str, recipient: str, sender_name: str, profile: StyleProfile
) -> str:
    """Generate email body following style guidelines."""
    # Extract recipient name from email if needed
    recipient_name = recipient.split("@")[0] if "@" in recipient else recipient

    # Handle None sender_name
    sender_name = sender_name or "[Your name]"

    # Use template if available
    if profile.example_template:
        body = profile.example_template.format(
            greeting=profile.greeting or "Hello",
            recipient=recipient_name,
            title="",  # Could be enhanced to detect titles
            content=message_content,
            signoff=profile.signoff or "Best regards,",
            sender=sender_name,
        )
    else:
        # Fallback body generation
        greeting = profile.greeting or "Hello"
        signoff = profile.signoff or "Best regards,"
        body = f"{greeting} {recipient_name},\n\n{message_content}\n\n{signoff}\n{sender_name}"

    return body


def _save_draft_to_file(draft: dict, file_path: str) -> None:
    """Save draft in email format for use with send(draft_file=...)."""
    from pathlib import Path

    Path(file_path).parent.mkdir(parents=True, exist_ok=True)

    email_content = f"""To: {draft["recipient"]}
Subject: {draft["subject"]}

{draft["body"]}"""

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(email_content)


# Style discovery tool
@mcp.tool(
    description="Get available email styles and their guidelines. Use before composing emails to understand available style prompts."
)
def get_email_styles() -> dict:
    """Return available email style prompts and their guidelines."""
    styles_info = {}
    for name, profile in STYLE_CONFIG.styles.items():
        styles_info[name] = {
            "tone": profile.tone,
            "guidelines": profile.guidelines[:5],  # First 5 guidelines
            "greeting": profile.greeting,
            "signoff": profile.signoff,
            "max_subject_len": profile.max_subject_len,
            "prompt_name": f"{name}_email",
            "example": create_example_body(profile, "colleague"),
        }

    return {
        "default_style": STYLE_CONFIG.default_style,
        "available_styles": styles_info,
        "usage": "Use the appropriate prompt (e.g., 'professional_email') before calling send() to compose emails in that style.",
    }
