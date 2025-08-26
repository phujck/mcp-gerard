"""Tests for email style configuration and prompts."""

import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import yaml

from mcp_handley_lab.email.style_config import (
    DEFAULT_CONFIG,
    EmailStyleConfig,
    StyleProfile,
    create_example_body,
    format_guidelines,
    get_style_prompt_messages,
    load_style_config,
)


class TestStyleProfile:
    """Test StyleProfile model."""

    def test_default_values(self):
        """Test default values for StyleProfile."""
        profile = StyleProfile()
        assert profile.tone == "professional"
        assert profile.system_message == "Compose professional emails"
        assert profile.guidelines == []
        assert profile.greeting is None
        assert profile.signoff is None
        assert profile.max_subject_len is None

    def test_custom_values(self):
        """Test custom values for StyleProfile."""
        profile = StyleProfile(
            tone="casual",
            system_message="Be friendly",
            guidelines=["Keep it short", "Use contractions"],
            greeting="Hi",
            signoff="Thanks,",
            max_subject_len=50,
        )
        assert profile.tone == "casual"
        assert profile.system_message == "Be friendly"
        assert len(profile.guidelines) == 2
        assert profile.greeting == "Hi"
        assert profile.signoff == "Thanks,"
        assert profile.max_subject_len == 50


class TestEmailStyleConfig:
    """Test EmailStyleConfig model."""

    def test_default_config(self):
        """Test default configuration."""
        config = EmailStyleConfig()
        assert config.version == "1"
        assert config.default_style == "professional"
        assert config.styles == {}

    def test_with_styles(self):
        """Test configuration with styles."""
        config = EmailStyleConfig(
            version="1",
            default_style="casual",
            styles={
                "casual": StyleProfile(tone="casual", greeting="Hi"),
                "formal": StyleProfile(tone="formal", greeting="Dear"),
            },
        )
        assert config.default_style == "casual"
        assert len(config.styles) == 2
        assert config.styles["casual"].tone == "casual"
        assert config.styles["formal"].tone == "formal"


class TestLoadStyleConfig:
    """Test configuration loading."""

    def test_load_default_config(self):
        """Test loading default config when no file exists."""
        with (
            patch.dict(os.environ, {}, clear=True),
            patch("pathlib.Path.exists", return_value=False),
        ):
            config = load_style_config()
            assert config == DEFAULT_CONFIG

    def test_load_from_environment_variable(self):
        """Test loading config from EMAIL_STYLE_FILE."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yml", delete=False) as f:
            config_data = {
                "version": "1",
                "default_style": "test",
                "styles": {
                    "test": {
                        "tone": "testing",
                        "guidelines": ["Test guideline"],
                    }
                },
            }
            yaml.dump(config_data, f)
            f.flush()

            try:
                with patch.dict(os.environ, {"EMAIL_STYLE_FILE": f.name}):
                    config = load_style_config()
                    assert config.default_style == "test"
                    assert config.styles["test"].tone == "testing"
            finally:
                os.unlink(f.name)

    def test_load_from_xdg_config(self):
        """Test loading from XDG_CONFIG_HOME."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir) / "mcp-email"
            config_dir.mkdir(parents=True)
            config_file = config_dir / "style.yml"

            config_data = {
                "version": "1",
                "default_style": "xdg_test",
                "styles": {"xdg_test": {"tone": "xdg"}},
            }
            with open(config_file, "w") as f:
                yaml.dump(config_data, f)

            with patch.dict(
                os.environ, {"XDG_CONFIG_HOME": tmpdir, "EMAIL_STYLE_FILE": ""}
            ):
                config = load_style_config()
                assert config.default_style == "xdg_test"

    def test_invalid_config_falls_back(self):
        """Test that invalid config falls back to default."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".yml", delete=False) as f:
            f.write("invalid: [yaml content")
            f.flush()

            try:
                with patch.dict(os.environ, {"EMAIL_STYLE_FILE": f.name}):
                    config = load_style_config()
                    assert config == DEFAULT_CONFIG
            finally:
                os.unlink(f.name)


class TestFormatGuidelines:
    """Test guideline formatting."""

    def test_format_empty_guidelines(self):
        """Test formatting empty guidelines."""
        assert format_guidelines([]) == ""

    def test_format_single_guideline(self):
        """Test formatting single guideline."""
        result = format_guidelines(["Keep it short"])
        assert result == "Keep it short"

    def test_format_multiple_guidelines(self):
        """Test formatting multiple guidelines."""
        guidelines = ["First rule", "Second rule", "Third rule"]
        result = format_guidelines(guidelines)
        assert result == "First rule; Second rule; Third rule"

    def test_format_truncates_long_text(self):
        """Test that long guidelines are truncated."""
        guidelines = ["Very " * 100 + "long guideline"]
        result = format_guidelines(guidelines, max_chars=50)
        assert len(result) == 50
        assert result.endswith("...")

    def test_format_normalizes_whitespace(self):
        """Test whitespace normalization."""
        guidelines = ["Has  multiple   spaces", "Has\n\nnewlines"]
        result = format_guidelines(guidelines)
        assert "  " not in result
        assert "\n" not in result


class TestCreateExampleBody:
    """Test example body creation."""

    def test_create_with_template(self):
        """Test creating example with template."""
        profile = StyleProfile(
            greeting="Hi",
            signoff="Thanks,",
            example_template="{greeting} {recipient},\n\nMessage: {content}\n\n{signoff}\n{sender}",
        )
        result = create_example_body(profile, "Alice")
        assert "Hi Alice," in result
        assert "Thanks," in result
        assert "[Your message here]" in result

    def test_create_without_template(self):
        """Test creating example without template."""
        profile = StyleProfile(greeting="Hello", signoff="Best,")
        result = create_example_body(profile, "Bob")
        assert "Hello Bob," in result
        assert "Best," in result
        assert "[Your message here]" in result

    def test_create_with_defaults(self):
        """Test creating example with default values."""
        profile = StyleProfile()
        result = create_example_body(profile, "User")
        assert "Hello User," in result
        assert "Best regards," in result


class TestGetStylePromptMessages:
    """Test prompt message generation."""

    def test_get_prompt_for_existing_style(self):
        """Test getting prompt messages for existing style."""
        config = EmailStyleConfig(
            default_style="test",
            styles={
                "test": StyleProfile(
                    tone="testing",
                    system_message="Test system message",
                    guidelines=["Guideline 1", "Guideline 2"],
                )
            },
        )
        messages = get_style_prompt_messages(
            "test", "Test content", "recipient@test.com", config
        )

        assert len(messages) == 2
        assert messages[0]["role"] == "system"
        assert "Test system message" in messages[0]["content"]
        assert "Guideline 1; Guideline 2" in messages[0]["content"]
        assert messages[1]["role"] == "user"
        assert "recipient@test.com" in messages[1]["content"]
        assert "Test content" in messages[1]["content"]

    def test_get_prompt_for_nonexistent_style(self):
        """Test getting prompt for style that doesn't exist."""
        config = EmailStyleConfig(
            default_style="fallback",
            styles={
                "fallback": StyleProfile(
                    tone="default", system_message="Default message"
                )
            },
        )
        messages = get_style_prompt_messages(
            "nonexistent", "Content", "user@test.com", config
        )

        assert len(messages) == 2
        assert "Default message" in messages[0]["content"]

    def test_get_prompt_with_default_config(self):
        """Test getting prompt with default configuration."""
        messages = get_style_prompt_messages(
            "professional", "Meeting request", "manager@company.com", DEFAULT_CONFIG
        )

        assert len(messages) == 2
        assert messages[0]["role"] == "system"
        assert "professional" in messages[0]["content"].lower()
        assert messages[1]["role"] == "user"
        assert "manager@company.com" in messages[1]["content"]
