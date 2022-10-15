from __future__ import annotations

from excelsigncheck.types import CertificateDetail, SignatureDetail


class ExcelSignature:
    """
    SignatureInfo is from https://learn.microsoft.com/pt-br/office/vba/api/overview/library-reference/signatureinfo-members-office
    """
    signature_info = None

    def __init__(self, signature_info):
        self.signature_info = signature_info

    @property
    def details(self) -> object:
        return self.signature_info.Details

    @property
    def certificate_verification_results(self) -> int:
        return self.details.CertificateVerificationResults

    @property
    def content_verification_results(self) -> int:
        return self.details.ContentVerificationResults

    @property
    def creator(self) -> int:
        return self.details.Creator

    @property
    def is_certificate_expired(self) -> bool:
        return self.details.IsCertificateExpired

    @property
    def is_certificate_revoked(self) -> bool:
        return self.details.IsCertificateRevoked

    @property
    def is_certificate_untrusted(self) -> bool:
        return self.details.IsCertificateUntrusted

    @property
    def is_valid(self) -> bool:
        return self.details.IsValid

    @property
    def read_only(self) -> bool:
        return self.details.ReadOnly

    @property
    def signature_comment(self) -> str:
        return self.details.SignatureComment

    @property
    def signature_image(self):
        return self.details.SignatureImage

    @property
    def signature_provider(self) -> str:
        return self.details.SignatureProvider

    @property
    def signature_text(self) -> str:
        return self.details.SignatureText

    def get_certificate_detail(self, detail_code: CertificateDetail | int):
        assert isinstance(detail_code, CertificateDetail) or isinstance(detail_code, int)
        return self.details.GetCertificateDetail(detail_code.value)

    def get_signature_detail(self,detail_code : SignatureDetail | int):
        assert isinstance(detail_code, SignatureDetail) or isinstance(detail_code, int)
        return self.details.GetSignatureDetail(detail_code.value)