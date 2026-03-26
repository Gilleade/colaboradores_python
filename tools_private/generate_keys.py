from pathlib import Path

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    keys_dir = base_dir / "keys"
    keys_dir.mkdir(parents=True, exist_ok=True)

    private_key_path = keys_dir / "private_key.pem"
    public_key_path = keys_dir / "public_key.pem"

    if private_key_path.exists() or public_key_path.exists():
        print("As chaves já existem.")
        print(f"Private key: {private_key_path}")
        print(f"Public key: {public_key_path}")
        return

    private_key = Ed25519PrivateKey.generate()
    public_key = private_key.public_key()

    private_bytes = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )

    public_bytes = public_key.public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo,
    )

    private_key_path.write_bytes(private_bytes)
    public_key_path.write_bytes(public_bytes)

    print("Chaves geradas com sucesso.")
    print(f"Private key: {private_key_path}")
    print(f"Public key: {public_key_path}")
    print("IMPORTANTE: nunca envie a private_key.pem para o cliente.")


if __name__ == "__main__":
    main()