class OfficeExtractorError(Exception):
    pass


class FileTypeError(OfficeExtractorError):
    pass


class NotAValidFileError(OfficeExtractorError):
    pass
