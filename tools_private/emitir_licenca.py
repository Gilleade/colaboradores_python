import argparse
import base64
import json
from datetime import datetime
from pathlib import Path

from cryptography.hazmat.primitives import serialization


DATE_FORMAT = "%Y-%m-%d"


def load_activation_request(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de ativação não encontrado: {path}")

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON inválido no arquivo de ativação: {path}") from exc

    required_fields = ["app_id", "app_name", "customer_type", "customer_name", "machine_id"]
    missing = [field for field in required_fields if field not in data]
    if missing:
        raise ValueError(
            f"Campos obrigatórios ausentes no arquivo de ativação: {', '.join(missing)}"
        )

    return data


def build_payload(args: argparse.Namespace) -> dict:
    
    if not args.customer_type:
        raise ValueError("É necessário informar customer_type ou usar --activation-request.")

    if not args.customer_name:
        raise ValueError("É necessário informar customer_name ou usar --activation-request.")
    
    payload = {
        "app_id": "colaboradores",
        "app_name": "Separador de Listas e Provas",
        "license_id": args.license_id,
        "customer_type": args.customer_type,
        "customer_name": args.customer_name,
        "license_type": args.license_type,
        "issued_at": datetime.now().strftime(DATE_FORMAT),
        "expires_at": None,
        "machine_id": args.machine_id,
        "notes": args.notes or "",
    }

    if args.license_type == "temporary":
        if not args.expires_at:
            raise ValueError("Licenças temporárias exigem --expires-at no formato YYYY-MM-DD.")
        datetime.strptime(args.expires_at, DATE_FORMAT)
        payload["expires_at"] = args.expires_at

    return payload


def sign_payload(payload: dict, private_key_path: Path) -> str:
    private_key = serialization.load_pem_private_key(
        private_key_path.read_bytes(),
        password=None,
    )

    canonical_json = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")

    signature = private_key.sign(canonical_json)
    return base64.b64encode(signature).decode("utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Emite uma licença assinada para o projeto.")
    parser.add_argument("--license-id", required=True)
    parser.add_argument("--customer-type", required=False, choices=["person", "sector", "company"])
    parser.add_argument("--customer-name", required=False, default="")
    parser.add_argument("--license-type", required=True, choices=["temporary", "perpetual"])
    parser.add_argument("--expires-at", required=False, help="Formato YYYY-MM-DD")
    parser.add_argument("--machine-id", required=False, default=None)
    parser.add_argument("--notes", required=False, default="")
    parser.add_argument(
        "--output",
        required=False,
        default="license.json",
        help="Caminho do arquivo de saída",
    )
    parser.add_argument(
        "--activation-request",
        required=False,
        help="Caminho para um activation_request.json gerado pelo programa",
    )

    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    private_key_path = base_dir / "keys" / "private_key.pem"

    if not private_key_path.exists():
        raise FileNotFoundError(f"Chave privada não encontrada em: {private_key_path}")

    if args.activation_request:
        activation_request_path = Path(args.activation_request).resolve()
        activation_data = load_activation_request(activation_request_path)

        if not args.customer_name:
            args.customer_name = activation_data.get("customer_name", "")

        if not args.customer_type:
            args.customer_type = activation_data.get("customer_type", "person")

        if not args.machine_id:
            args.machine_id = activation_data.get("machine_id")

    payload = build_payload(args)
    signature = sign_payload(payload, private_key_path)

    license_data = {
        **payload,
        "signature": signature,
    }

    output_path = Path(args.output).resolve()
    output_path.write_text(
        json.dumps(license_data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("Licença gerada com sucesso.")
    print(f"Arquivo: {output_path}")


if __name__ == "__main__":
    main()