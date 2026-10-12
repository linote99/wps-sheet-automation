from __future__ import annotations

import base64
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

import licensing
from cryptography.hazmat.primitives import serialization


class LicenseCryptoTests(unittest.TestCase):
    machine = "WPSA-01234567-89ABCDEF-01234567-89ABCDEF-01234567"
    other_machine = "WPSA-11234567-89ABCDEF-01234567-89ABCDEF-01234567"
    private_key = (Path(__file__).resolve().parent / "licensing_private" / "issuer_private_key.pem").read_bytes()

    def test_machine_code_is_stable_and_grouped(self):
        self.assertEqual(licensing.machine_code("test-guid"), licensing.machine_code("test-guid"))
        self.assertRegex(licensing.machine_code("test-guid"), r"^WPSA-(?:[0-9A-F]{8}-){4}[0-9A-F]{8}$")
        self.assertEqual(licensing.normalize_machine_code(self.machine.replace("-", "").lower()), self.machine)

    def test_issuer_private_key_matches_program_public_key(self):
        private_key = serialization.load_pem_private_key(self.private_key, password=None)
        public_der = private_key.public_key().public_bytes(
            serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo,
        )
        encoded = base64.urlsafe_b64encode(public_der).decode("ascii").rstrip("=")
        self.assertEqual(encoded, licensing.PUBLIC_KEY_B64)

    def test_issued_token_binds_machine_and_inclusive_expiry(self):
        expiry = date(2030, 2, 28)
        token = licensing.generate_license(self.machine, expiry, self.private_key)
        self.assertEqual(licensing.verify_license(token, self.machine, today=expiry), expiry)
        with self.assertRaisesRegex(licensing.LicensingError, "另一台电脑"):
            licensing.verify_license(token, self.other_machine, today=date(2029, 1, 1))
        with self.assertRaisesRegex(licensing.LicensingError, "到期"):
            licensing.verify_license(token, self.machine, today=expiry + timedelta(days=1))

    def test_expiry_inspection_reports_expired_date_without_authorizing_expired_key(self):
        expiry = date(2030, 2, 28)
        token = licensing.generate_license(self.machine, expiry, self.private_key)
        self.assertEqual(expiry, licensing.inspect_license_expiry(token, self.machine))
        with self.assertRaisesRegex(licensing.LicensingError, "到期"):
            licensing.verify_license(token, self.machine, today=expiry + timedelta(days=1))

    def test_modified_payload_or_signature_is_rejected(self):
        token = licensing.generate_license(self.machine, "2030-12-31", self.private_key)
        payload_part, signature_part = token.split(".")
        changed_payload = base64.urlsafe_b64encode(b'{"app":"tampered"}').decode().rstrip("=")
        with self.assertRaises(licensing.LicensingError):
            licensing.verify_license(changed_payload + "." + signature_part, self.machine, today=date(2030, 1, 1))
        changed_signature = ("A" if signature_part[0] != "A" else "B") + signature_part[1:]
        with self.assertRaises(licensing.LicensingError):
            licensing.verify_license(payload_part + "." + changed_signature, self.machine, today=date(2030, 1, 1))

    def test_invalid_private_key_or_machine_code_is_rejected(self):
        with self.assertRaises(licensing.LicensingError):
            licensing.generate_license("not-a-machine-code", "2030-01-01", self.private_key)
        with self.assertRaises(licensing.LicensingError):
            licensing.generate_license(self.machine, "not-a-date", self.private_key)
        with self.assertRaises(licensing.LicensingError):
            licensing.generate_license(self.machine, "2030-01-01", b"not a pem")


class TrialStateTests(unittest.TestCase):
    machine = LicenseCryptoTests.machine

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        base = Path(self.temp.name)
        self.state = base / "trial.dat"
        self.license_file = base / "license.key"

    def tearDown(self):
        self.temp.cleanup()

    def test_trial_is_seven_calendar_days_from_first_launch(self):
        start = date(2026, 10, 7)
        first = licensing.evaluate_authorization(machine=self.machine, today=start,
                                                 state_path=self.state, license_path=self.license_file)
        self.assertEqual(first.state, "trial")
        self.assertEqual(first.trial_remaining_days, 7)
        sixth = licensing.evaluate_authorization(machine=self.machine, today=start + timedelta(days=6),
                                                 state_path=self.state, license_path=self.license_file)
        self.assertEqual(sixth.state, "trial")
        self.assertEqual(sixth.trial_remaining_days, 1)
        expired = licensing.evaluate_authorization(machine=self.machine, today=start + timedelta(days=7),
                                                   state_path=self.state, license_path=self.license_file)
        self.assertEqual(expired.state, "expired")

    def test_system_clock_rollback_blocks_trial(self):
        start = date(2026, 10, 7)
        licensing.evaluate_authorization(machine=self.machine, today=start + timedelta(days=2),
                                         state_path=self.state, license_path=self.license_file)
        rollback = licensing.evaluate_authorization(machine=self.machine, today=start + timedelta(days=1),
                                                    state_path=self.state, license_path=self.license_file)
        self.assertEqual(rollback.state, "clock_rollback")

    def test_trial_state_is_machine_bound_and_corruption_fails_closed(self):
        today = date(2026, 10, 7)
        licensing.evaluate_authorization(machine=self.machine, today=today,
                                         state_path=self.state, license_path=self.license_file)
        copied = licensing.evaluate_authorization(machine=LicenseCryptoTests.other_machine, today=today,
                                                  state_path=self.state, license_path=self.license_file)
        self.assertEqual(copied.state, "state_error")
        self.state.write_bytes(b"broken")
        corrupt = licensing.evaluate_authorization(machine=self.machine, today=today,
                                                   state_path=self.state, license_path=self.license_file)
        self.assertEqual(corrupt.state, "state_error")

    def test_installed_license_authorizes_even_after_trial(self):
        start = date(2026, 10, 7)
        state = self.state
        licensing.evaluate_authorization(machine=self.machine, today=start,
                                         state_path=state, license_path=self.license_file)
        token = licensing.generate_license(
            self.machine, "2027-10-07", LicenseCryptoTests.private_key,
        )
        expiry = licensing.install_license(token, machine=self.machine, license_path=self.license_file, today=start)
        self.assertEqual(expiry, date(2027, 10, 7))
        status = licensing.evaluate_authorization(machine=self.machine, today=date(2027, 1, 1),
                                                  state_path=state, license_path=self.license_file)
        self.assertEqual(status.state, "licensed")


class DistributionPackagingTests(unittest.TestCase):
    project = Path(__file__).resolve().parent

    def test_windows_builder_builds_both_apps_and_copies_plugin_and_admin_key(self):
        script = (self.project / "install_and_build.bat").read_text(encoding="utf-8").lower()
        self.assertIn("wpssheetautomation_keygen", script)
        self.assertIn("issuer_private_key.pem", script)
        self.assertIn("ufitools_keygen_plugin.html", script)
        self.assertIn("--collect-all cryptography", script)

    def test_ufi_plugin_is_offline_and_includes_bundled_signer(self):
        html = (self.project / "ufitools_keygen_plugin.html").read_text(encoding="utf-8")
        self.assertEqual(html.lower().count("<script>"), 2)
        self.assertNotIn("<script src=", html.lower())
        self.assertNotIn("fetch(", html)
        self.assertNotIn("XMLHttpRequest", html)
        self.assertIn("WpsLicenseVendor.elliptic", html)
        self.assertIn("privateScalarFromPem", html)


if __name__ == "__main__":
    unittest.main(verbosity=2)
