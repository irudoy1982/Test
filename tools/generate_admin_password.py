from __future__ import annotations

import base64
import getpass
import hashlib
import secrets


PBKDF2_ITERATIONS = 600_000


def main() -> None:
    password = getpass.getpass("New admin password: ")
    confirmation = getpass.getpass("Repeat admin password: ")
    if not password:
        raise SystemExit("Password cannot be empty.")
    if password != confirmation:
        raise SystemExit("Passwords do not match.")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
    )
    salt_text = base64.urlsafe_b64encode(salt).decode("ascii")
    digest_text = base64.urlsafe_b64encode(digest).decode("ascii")
    print(f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt_text}${digest_text}")


if __name__ == "__main__":
    main()
