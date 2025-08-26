"""Email style configuration management."""

import os
import re
from pathlib import Path

import yaml
from pydantic import BaseModel, Field

# Configuration search paths in order of precedence
CONFIG_SEARCH_PATHS = [
    lambda: os.environ.get("EMAIL_STYLE_FILE"),
    lambda: os.path.join(
        os.environ.get("XDG_CONFIG_HOME", os.path.expanduser("~/.config")),
        "mcp-email",
        "style.yml",
    ),
    lambda: os.path.expanduser("~/.emailstyle.yml"),
]


class StyleProfile(BaseModel):
    """Email style profile configuration."""

    tone: str = Field(default="professional", description="The tone of the email style")
    system_message: str = Field(
        default="Compose professional emails",
        description="System message for LLM guidance",
    )
    guidelines: list[str] = Field(
        default_factory=list,
        description="List of style guidelines (3-6 bullets recommended)",
    )
    greeting: str | None = Field(
        default=None, description="Default greeting (e.g., 'Hello', 'Hi')"
    )
    signoff: str | None = Field(
        default=None, description="Default signoff (e.g., 'Best regards,', 'Thanks,')"
    )
    max_subject_len: int | None = Field(
        default=None, description="Maximum subject line length"
    )
    example_template: str | None = Field(
        default=None, description="Example email template for this style"
    )


class EmailStyleConfig(BaseModel):
    """Root email style configuration."""

    version: str = Field(default="1", description="Configuration version")
    default_style: str = Field(
        default="professional", description="Default style to use"
    )
    styles: dict[str, StyleProfile] = Field(
        default_factory=dict, description="Available email styles"
    )


# Default fallback configuration
DEFAULT_CONFIG = EmailStyleConfig(
    version="1",
    default_style="professional",
    styles={
        "professional": StyleProfile(
            tone="professional",
            system_message="Compose professional business emails with clear structure and formal tone",
            guidelines=[
                "Keep subject line concise and descriptive (max 65 characters)",
                "Use formal salutation and closing",
                "Structure content with clear paragraphs",
                "Include specific action items and deadlines",
                "Maintain professional tone throughout",
            ],
            greeting="Hello",
            signoff="Best regards,",
            max_subject_len=65,
            example_template="{greeting} {recipient},\n\nI hope this email finds you well. {content}\n\n{signoff}\n{sender}",
        ),
        "casual": StyleProfile(
            tone="casual",
            system_message="Compose friendly, informal emails with a relaxed tone",
            guidelines=[
                "Use friendly, conversational language",
                "Keep it brief and to the point",
                "Contractions and informal phrases are fine",
                "Use simple, direct sentences",
            ],
            greeting="Hi",
            signoff="Thanks,",
            example_template="Hi {recipient},\n\n{content}\n\nThanks!\n{sender}",
        ),
        "academic": StyleProfile(
            tone="academic",
            system_message="Compose formal academic emails with scholarly tone and precise language",
            guidelines=[
                "Use formal academic language and proper titles",
                "Be precise and specific in references",
                "Maintain scholarly tone and objectivity",
                "Include relevant citations or references when applicable",
                "Use complete sentences and proper grammar",
            ],
            greeting="Dear",
            signoff="Sincerely,",
            max_subject_len=80,
            example_template="Dear {title} {recipient},\n\n{content}\n\n{signoff}\n{sender}",
        ),
    },
)


def load_style_config() -> EmailStyleConfig:
    """Load email style configuration from available sources."""
    for path_resolver in CONFIG_SEARCH_PATHS:
        path = path_resolver()
        if not path:
            continue

        config_path = Path(path)
        if config_path.exists():
            try:
                with open(config_path, encoding="utf-8") as f:
                    data = yaml.safe_load(f) or {}
                return EmailStyleConfig(**data)
            except Exception as e:
                # Log error but continue to next path or fallback
                print(f"Error loading config from {config_path}: {e}")
                continue

    return DEFAULT_CONFIG


def format_guidelines(guidelines: list[str], max_chars: int = 360) -> str:
    """Format guidelines into a concise string for tool descriptions."""
    if not guidelines:
        return ""

    # Take up to 6 guidelines and join with semicolons
    text = "; ".join(guidelines[:6])
    # Normalize whitespace
    text = re.sub(r"\s+", " ", text).strip()
    # Truncate if too long
    if len(text) > max_chars:
        text = text[: max_chars - 3] + "..."
    return text


def create_example_body(profile: StyleProfile, recipient: str = "John") -> str:
    """Create an example email body for the given style."""
    if profile.example_template:
        return profile.example_template.format(
            greeting=profile.greeting or "Hello",
            recipient=recipient,
            title="Dr.",
            content="[Your message here]",
            signoff=profile.signoff or "Best regards,",
            sender="[Your name]",
        )

    # Fallback to simple example
    greeting = profile.greeting or "Hello"
    signoff = profile.signoff or "Best regards,"
    return f"{greeting} {recipient},\n\n[Your message here]\n\n{signoff}\n[Your name]"


def get_style_prompt_messages(
    style_name: str, message_content: str, recipient: str, config: EmailStyleConfig
) -> list[dict[str, str]]:
    """Generate prompt messages for a specific email style."""
    profile = config.styles.get(style_name)
    if not profile:
        profile = config.styles.get(
            config.default_style, DEFAULT_CONFIG.styles["professional"]
        )

    guidelines_text = format_guidelines(profile.guidelines)

    return [
        {
            "role": "system",
            "content": f"{profile.system_message}. Guidelines: {guidelines_text}",
        },
        {
            "role": "user",
            "content": f"Compose an email to {recipient} about: {message_content}",
        },
    ]


# Load configuration once at module import
STYLE_CONFIG = load_style_config()
