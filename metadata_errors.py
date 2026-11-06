class MetadataError(ValueError):
    """A selected file cannot be safely handled by the metadata editor."""


class EncryptedFileError(MetadataError):
    """Password-to-open encryption is outside the supported scope."""
