class OfficeExtractorError(Exception):
    pass


class FileTypeError(OfficeExtractorError):
    pass


class NotAValidZipFileError(OfficeExtractorError):
    pass
