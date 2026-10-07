import os
from pathlib import Path

from dotenv import load_dotenv
from bale import auth


ENV_FILE = Path(".env")


def normalize_phone(value: str) -> int:
    value = value.strip().replace("+", "").replace(" ", "").replace("-", "")
    if value.startswith("0"):
        value = "98" + value[1:]
    if not value.isdigit():
        raise ValueError("Invalid mobile number.")
    return int(value)


def upsert_env(key: str, value: str):
    lines = []
    if ENV_FILE.exists():
        lines = ENV_FILE.read_text(encoding="utf-8").splitlines()

    out = []
    found = False
    for line in lines:
        if line.startswith(key + "="):
            out.append(f"{key}={value}")
            found = True
        else:
            out.append(line)
    if not found:
        out.append(f"{key}={value}")

    ENV_FILE.write_text("\n".join(out).rstrip() + "\n", encoding="utf-8")


def main():
    load_dotenv()
    print("Bale account login with OTP")
    phone = normalize_phone(input("Bale mobile number (example: 0912... or +98912...): "))

    session = auth.start_phone_auth(phone)
    print("Login code sent.")
    code = input("OTP code: ").strip()

    result = auth.validate_code(session.transaction_hash, code)
    upsert_env("BALE_TOKEN", result.access_token)

    print("Login successful.")
    print("The token was saved to the .env file. Do not commit it to GitHub.")


if __name__ == "__main__":
    main()
