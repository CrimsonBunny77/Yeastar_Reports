"""User-facing exceptions. Messages are safe to display directly."""


class ReportError(Exception):
    """Base class for errors with a friendly message."""


class EmptyUploadError(ReportError):
    pass


class InvalidFileError(ReportError):
    pass


class MissingColumnsError(ReportError):
    pass


class MappingError(ReportError):
    pass


class NoTransfersError(ReportError):
    pass
