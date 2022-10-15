import pathlib
import unittest
from excelsigncheck import ExcelSignCheck
from excelsigncheck.types import CertificateDetail, SignatureDetail


class ExcelSignCheckTestCase(unittest.TestCase):

    def setUp(self) -> None:
        ROOT = pathlib.Path(__file__).parent.absolute()
        self.file_with_signature = str(ROOT.joinpath('file-with-signature.xlsx'))
        self.file_without_signature = str(ROOT.joinpath('file-without-signature.xlsx'))

    def test_signatures_iter(self):
        sign_check = ExcelSignCheck(self.file_with_signature)
        sign_check.open()
        num_signatures = len(list(sign_check.signatures))
        self.assertGreater(num_signatures, 0)
        sign_check.close()

        sign_check = ExcelSignCheck(self.file_without_signature)
        sign_check.open()
        num_signatures = len(list(sign_check.signatures))
        self.assertEqual(num_signatures, 0)
        sign_check.close()

    def test_certificate_details(self):
        with ExcelSignCheck(self.file_with_signature) as handler:
            for signature in handler.signatures:
                for certdet in CertificateDetail:
                    value = signature.get_certificate_detail(certdet)
                    print('%s: %s' % (certdet.name, value))

    def test_signature_details(self):
        with ExcelSignCheck(self.file_with_signature) as handler:
            for signature in handler.signatures:
                for signdet in SignatureDetail:

                    if signdet.value == 14:
                        # This detail raises pywintypes.com_error
                        continue

                    value = signature.get_signature_detail(signdet)
                    print('%s: %s' % (signdet.name, value))

if __name__ == '__main__':
    unittest.main()
