"""Tests for demo_agent.config — DemoConfig validation and defaults."""

from __future__ import annotations

import os

import pytest

from demo_agent.config import DemoConfig


class TestDemoConfigDefaults:

    def test_default_d365_base_url(self):
        config = DemoConfig()
        assert "crm.dynamics.com" in config.d365_base_url

    def test_default_llm_model(self):
        config = DemoConfig()
        assert config.llm_model == "gpt-4o"

    def test_default_headless_false(self):
        config = DemoConfig()
        assert config.headless is False

    def test_default_voice_disabled(self):
        config = DemoConfig()
        assert config.voice_enabled is False

    def test_default_slow_mo(self):
        config = DemoConfig()
        assert config.slow_mo == 50

    def test_default_max_demo_steps(self):
        config = DemoConfig()
        assert config.max_demo_steps == 20


class TestDemoConfigCredentialProperties:

    def test_use_dataverse_auth_false_by_default(self):
        config = DemoConfig()
        assert config.use_dataverse_auth is False

    def test_use_dataverse_auth_true(self, monkeypatch):
        monkeypatch.setenv("DATAVERSE_TENANT_ID", "tenant-123")
        monkeypatch.setenv("DATAVERSE_CLIENT_ID", "client-123")
        monkeypatch.setenv("DATAVERSE_CLIENT_SECRET", "secret-123")
        config = DemoConfig()
        assert config.use_dataverse_auth is True

    def test_use_dataverse_auth_partial_false(self, monkeypatch):
        monkeypatch.setenv("DATAVERSE_TENANT_ID", "tenant-123")
        # Missing client_id and client_secret
        config = DemoConfig()
        assert config.use_dataverse_auth is False

    def test_use_azure_openai_false_by_default(self):
        config = DemoConfig()
        assert config.use_azure_openai is False

    def test_use_azure_openai_true(self, monkeypatch):
        monkeypatch.setenv("AZURE_OPENAI_ENDPOINT", "https://my.openai.azure.com")
        monkeypatch.setenv("AZURE_OPENAI_API_KEY", "key-123")
        config = DemoConfig()
        assert config.use_azure_openai is True

    def test_use_github_copilot_false_by_default(self):
        config = DemoConfig()
        assert config.use_github_copilot is False

    def test_use_github_copilot_true(self, monkeypatch):
        monkeypatch.setenv("GITHUB_TOKEN", "ghp_abc123")
        config = DemoConfig()
        assert config.use_github_copilot is True

    def test_use_openai_false_by_default(self):
        config = DemoConfig()
        assert config.use_openai is False

    def test_use_openai_true(self, monkeypatch):
        monkeypatch.setenv("OPENAI_API_KEY", "sk-abc123")
        config = DemoConfig()
        assert config.use_openai is True


class TestDemoConfigValidation:

    def test_validate_no_llm_credentials(self):
        config = DemoConfig()
        errors = config.validate()
        assert len(errors) == 1
        assert "No LLM credentials" in errors[0]

    def test_validate_azure_openai_ok(self, monkeypatch):
        monkeypatch.setenv("AZURE_OPENAI_ENDPOINT", "https://my.openai.azure.com")
        monkeypatch.setenv("AZURE_OPENAI_API_KEY", "key-123")
        config = DemoConfig()
        assert config.validate() == []

    def test_validate_openai_ok(self, monkeypatch):
        monkeypatch.setenv("OPENAI_API_KEY", "sk-abc123")
        config = DemoConfig()
        assert config.validate() == []

    def test_validate_github_ok(self, monkeypatch):
        monkeypatch.setenv("GITHUB_TOKEN", "ghp_abc123")
        config = DemoConfig()
        assert config.validate() == []

    def test_validate_anthropic_ok(self, monkeypatch):
        monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-abc123")
        config = DemoConfig()
        assert config.validate() == []

    def test_env_overrides(self, monkeypatch):
        monkeypatch.setenv("D365_BASE_URL", "https://custom.crm.dynamics.com")
        monkeypatch.setenv("LLM_MODEL", "gpt-3.5-turbo")
        monkeypatch.setenv("BROWSER_HEADLESS", "true")
        monkeypatch.setenv("VOICE_ENABLED", "true")
        config = DemoConfig()
        assert config.d365_base_url == "https://custom.crm.dynamics.com"
        assert config.llm_model == "gpt-3.5-turbo"
        assert config.headless is True
        assert config.voice_enabled is True
