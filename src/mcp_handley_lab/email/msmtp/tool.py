"""MSMTP email sending provider with style support."""

from pathlib import Path

from pydantic import BaseModel, Field

from mcp_handley_lab.common.process import run_command
from mcp_handley_lab.email.common import mcp
from mcp_handley_lab.email.style_config import (
    STYLE_CONFIG,
    create_example_body,
    format_guidelines,
    get_style_prompt_messages,
)

# Generate dynamic descriptions from loaded style configuration
DEFAULT_STYLE = STYLE_CONFIG.styles.get(
    STYLE_CONFIG.default_style,
    list(STYLE_CONFIG.styles.values())[0] if STYLE_CONFIG.styles else None,
)
GUIDELINES = format_guidelines(DEFAULT_STYLE.guidelines) if DEFAULT_STYLE else ""
EXAMPLE = create_example_body(DEFAULT_STYLE) if DEFAULT_STYLE else ""

# Dynamic tool description with style guidance
SEND_DESCRIPTION = (
    f"Send email using msmtp. Default style: '{STYLE_CONFIG.default_style}'. "
    f"Before composing, use the appropriate email style prompt (e.g., 'professional_email', 'casual_email'). "
    f"Guidelines for {STYLE_CONFIG.default_style}: {GUIDELINES}"
)

# Dynamic field descriptions
SUBJECT_DESCRIPTION = "The subject line of the email. " + (
    f"Max {DEFAULT_STYLE.max_subject_len} characters."
    if DEFAULT_STYLE and DEFAULT_STYLE.max_subject_len
    else ""
)

BODY_DESCRIPTION = (
    f"The main content (body) of the email. Should follow the active style guidelines. "
    f"For '{STYLE_CONFIG.default_style}' style: {GUIDELINES}"
)


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
    to: str = Field(..., description="The primary recipient's email address."),
    subject: str = Field(..., description=SUBJECT_DESCRIPTION),
    body: str = Field(..., description=BODY_DESCRIPTION),
    account: str = Field(
        default="",
        description="The msmtp account to send from. If empty, the default account is used. Use 'list_accounts' to see options.",
    ),
    cc: str = Field(
        default="",
        description="Comma-separated list of email addresses for CC recipients.",
    ),
    bcc: str = Field(
        default="",
        description="Comma-separated list of email addresses for BCC recipients.",
    ),
) -> SendResult:
    """Send an email using msmtp with existing ~/.msmtprc configuration."""
    email_content = f"To: {to}\n"
    email_content += f"Subject: {subject}\n"

    if cc:
        email_content += f"Cc: {cc}\n"
    if bcc:
        email_content += f"Bcc: {bcc}\n"

    email_content += "\n"
    email_content += body

    cmd = ["msmtp"]
    if account:
        cmd.extend(["-a", account])

    recipients = [to]
    if cc:
        recipients.extend([addr.strip() for addr in cc.split(",")])
    if bcc:
        recipients.extend([addr.strip() for addr in bcc.split(",")])

    cmd.extend(recipients)

    input_bytes = email_content.encode()
    stdout, stderr = run_command(cmd, input_data=input_bytes)

    cc_list = [addr.strip() for addr in cc.split(",")] if cc else []
    bcc_list = [addr.strip() for addr in bcc.split(",")] if bcc else []

    return SendResult(
        recipient=to,
        account_used=account,
        cc_recipients=cc_list,
        bcc_recipients=bcc_list,
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


# MCP Prompts for email styles
@mcp.prompt()
def professional_email(
    message_content: str = Field(
        ..., description="The core content/purpose of the email"
    ),
    recipient: str = Field(..., description="The recipient's name or email"),
) -> list[dict[str, str]]:
    """Compose a professional business email with formal tone and clear structure."""
    return get_style_prompt_messages(
        "professional", message_content, recipient, STYLE_CONFIG
    )


@mcp.prompt()
def casual_email(
    message_content: str = Field(
        ..., description="The core content/purpose of the email"
    ),
    recipient: str = Field(..., description="The recipient's name or email"),
) -> list[dict[str, str]]:
    """Compose a casual, friendly email with relaxed tone."""
    return get_style_prompt_messages("casual", message_content, recipient, STYLE_CONFIG)


@mcp.prompt()
def academic_email(
    message_content: str = Field(
        ..., description="The core content/purpose of the email"
    ),
    recipient: str = Field(..., description="The recipient's name or email"),
) -> list[dict[str, str]]:
    """Compose a formal academic email with scholarly tone."""
    return get_style_prompt_messages(
        "academic", message_content, recipient, STYLE_CONFIG
    )


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
