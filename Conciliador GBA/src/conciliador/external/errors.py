from __future__ import annotations


class ExternalConfigError(ValueError):
    """Invalid API configuration (400)."""


class ExternalSchemaError(ValueError):
    """Invalid/missing critical fields from external API (424)."""


class ExternalProviderError(RuntimeError):
    """Provider call failed (502)."""

    def __init__(self, provider: str, message: str, *, status_code: int | None = None, request_id: str | None = None):
        super().__init__(message)
        self.provider = provider
        self.status_code = status_code
        self.request_id = request_id


class ExternalTimeoutError(ExternalProviderError):
    """Provider timeout (502)."""
