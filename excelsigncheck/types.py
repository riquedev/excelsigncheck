from enum import Enum


class CertificateDetail(Enum):
    """
        https://learn.microsoft.com/en-us/office/vba/api/office.certificatedetail
    """
    CERT_DET_AVAILABLE = 0  # Specifies that the digital certificate is available for signing.
    CERT_DET_EXPIRATION_DATE = 3  # The expiration date of the certificate.
    CERT_DET_ISSUER = 2  # The issuing authority of the certification.
    CERT_DET_SUBJECT = 1  # The holder of a Private Key corresponding to a Public Key.
    CERT_DET_THUMBPRINT = 4  # A hash of the certificate's complete contents.


class SignatureDetail(Enum):
    """
    https://learn.microsoft.com/pt-br/office/vba/api/office.signaturedetail
    """
    SIG_DET_APPLICATION_NAME = 1  # Specifies the application name.
    SIG_DET_APPLICATION_VERSION = 2  # Specifies the application version.
    SIG_DET_COLOR_DEPTH = 8  # Specifies the color depth.
    SIG_DET_DEL_SUGG_SIGNER = 16  # Specifies the suggested signer delegate.
    SIG_DET_DEL_SUGG_SIGNER_EMAIL = 20  # Specifies the suggested signer's delegate's email.
    SIG_DET_DEL_SUGG_EMAIL_SET = 21  # Indicates whether an email for a suggested signer delegate has been specified.
    SIG_DET_DEL_SUGG_SIGNER_LINE_2 = 18  # Specifies the suggested signer's delegate's signature line.
    SIG_DET_DEL_SUGG_SIGNER_LINE_2_SET = 19  # Specifies the set of suggested signer's delegate's signature lines.
    SIG_DET_DEL_SUGG_SIGNER_SET = 17  # Specifies the set of suggested signer's delegates.
    SIG_DET_DOC_PREVIEW_IMG = 10  # Specifies the document preview image.
    SIG_DET_HASH_ALGORITHM = 14  # Specifies the hash algorithm.
    SIG_DET_HORIZ_RESOLUTION = 6  # Specifies the horizontal resolution.
    SIG_DET_IP_CURRENT_VIEW = 12  # Specifies the IP current view.
    SIG_DET_IP_FORM_HASH = 11  # Specifies the IP form hash.
    SIG_DET_LOCAL_SIGNING_TIME = 0  # Specifies the local signing time.
    SIG_DET_NUMBERS_OF_MONITORS = 5  # Specifies the number of monitors.
    SIG_DET_OFFICE_VERSION = 3  # Specifies the Office version.
    SIG_DET_SHOULD_SHOW_VIEW_WARNING = 15  # Specifies the Should Show View Warning setting.
    SIG_DET_SIGNATURE_TYPE = 13  # Specifies the signed data.
    SIG_DET_VERT_RESOLUTION = 7  # Specifies the vertical resolution.
    SIG_DET_WINDOWS_VERSION = 4  # Specifies the Windows version.
