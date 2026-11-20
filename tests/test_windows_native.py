from hashlib import sha256
from io import BytesIO
import base64
import json
from pathlib import Path
import unittest
from unittest.mock import MagicMock, patch

from windows_local import EXTENSION_ORIGIN
from windows_native import MESSAGE_LIMIT, read_message, serve_native, write_message


ORIGIN = "https://conectores.example"
REQUEST = {"id": "7aa02a21-877d-4135-a219-2f0baea3692e", "command": "start", "origin": ORIGIN}


class NativeHostTests(unittest.TestCase):
    def test_public_manifest_key_matches_registered_native_extension(self):
        manifest = Path(__file__).resolve().parents[1] / "extension" / "manifest.json"
        public_key = base64.b64decode(json.loads(manifest.read_text(encoding="utf-8"))["key"])
        identifier = sha256(public_key).hexdigest()[:32].translate(
            str.maketrans("0123456789abcdef", "abcdefghijklmnop")
        )
        self.assertEqual(f"chrome-extension://{identifier}/", EXTENSION_ORIGIN)

    def test_messages_have_four_byte_lengths_and_enforce_a_size_limit(self):
        stream = BytesIO()
        write_message(stream, REQUEST)
        stream.seek(0)
        self.assertEqual(read_message(stream), REQUEST)
        self.assertIsNone(read_message(stream))
        with self.assertRaises(ValueError):
            read_message(BytesIO((MESSAGE_LIMIT + 1).to_bytes(4, "little")))

    def test_only_the_configured_origin_can_start_local_generator_once(self):
        inputs = BytesIO()
        for request in (
            {**REQUEST, "origin": "https://distinta.example"},
            {**REQUEST, "command": "other"},
            {**REQUEST, "filename": "unarchivo.dwg"},
            REQUEST,
            REQUEST,
        ):
            write_message(inputs, request)
        inputs.seek(0)
        output = BytesIO()
        server, pool = MagicMock(), MagicMock()
        with patch("windows_native.launcher_is_running", return_value=False), \
            patch("windows_native.get_local_server", return_value=(server, pool)) as launcher:
            serve_native(inputs, output, ORIGIN)
        launcher.assert_called_once()
        self.assertEqual(launcher.call_args.args[0], ORIGIN)
        server.run.assert_called_once()
        server.close.assert_called_once()
        pool.shutdown.assert_called_once_with(wait=True)
        output.seek(0)
        self.assertEqual([read_message(output)["ok"] for _ in range(5)], [False, False, False, True, True])


if __name__ == "__main__":
    unittest.main()
