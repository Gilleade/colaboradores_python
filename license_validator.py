import base64
import hashlib
import json
import platform
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import serialization


DATE_FORMAT = "%Y-%m-%d"
CLOCK_ROLLBACK_TOLERANCE_MINUTES = 5


class LicenseValidationError(Exception):
    pass


@dataclass
class LicenseValidationResult:
    license_id: str
    customer_name: str
    customer_type: str
    license_type: str
    expires_at: Optional[str]
    machine_bound: bool
    license_path: str


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_license_path(base_dir: Path) -> Path:
    return base_dir / "license.json"


def get_public_key_path(base_dir: Path) -> Path:
    return base_dir / "public_key.pem"


def get_state_path(base_dir: Path) -> Path:
    state_dir = base_dir / "colaboradores_data"
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir / "license_state.json"


def read_json_file(path: Path) -> dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise LicenseValidationError(f"Arquivo não encontrado: {path}")
    except json.JSONDecodeError as exc:
        raise LicenseValidationError(f"JSON inválido em {path}: {exc}") from exc


def build_canonical_payload(license_data: dict[str, Any]) -> bytes:
    payload = {k: v for k, v in license_data.items() if k != "signature"}
    return json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def verify_signature(license_data: dict[str, Any], public_key_path: Path) -> None:
    if not public_key_path.exists():
        raise LicenseValidationError(f"Chave pública não encontrada: {public_key_path}")

    signature_b64 = license_data.get("signature")
    if not signature_b64:
        raise LicenseValidationError("A licença não possui assinatura.")

    try:
        signature = base64.b64decode(signature_b64)
    except Exception as exc:
        raise LicenseValidationError("A assinatura da licença está inválida.") from exc

    public_key = serialization.load_pem_public_key(public_key_path.read_bytes())
    payload = build_canonical_payload(license_data)

    try:
        public_key.verify(signature, payload)
    except InvalidSignature as exc:
        raise LicenseValidationError("A assinatura da licença não confere.") from exc


def validate_required_fields(license_data: dict[str, Any]) -> None:
    required_fields = [
        "app_id",
        "app_name",
        "license_id",
        "customer_type",
        "customer_name",
        "license_type",
        "issued_at",
    ]

    missing = [field for field in required_fields if not license_data.get(field)]
    if missing:
        raise LicenseValidationError(
            f"Campos obrigatórios ausentes na licença: {', '.join(missing)}"
        )


def validate_expiration(license_data: dict[str, Any]) -> None:
    license_type = license_data.get("license_type")
    today = datetime.now().date()

    if license_type == "temporary":
        expires_at = license_data.get("expires_at")
        if not expires_at:
            raise LicenseValidationError("Licenças temporárias exigem expires_at.")

        try:
            expires_date = datetime.strptime(expires_at, DATE_FORMAT).date()
        except ValueError as exc:
            raise LicenseValidationError(
                "O campo expires_at está fora do formato YYYY-MM-DD."
            ) from exc

        if today > expires_date:
            raise LicenseValidationError(
                f"Licença expirada em {expires_at}."
            )

    elif license_type == "perpetual":
        return

    else:
        raise LicenseValidationError(
            f"Tipo de licença inválido: {license_type}"
        )


def get_windows_machine_guid() -> Optional[str]:
    if platform.system().lower() != "windows":
        return None

    try:
        import winreg

        access_flags = winreg.KEY_READ
        if hasattr(winreg, "KEY_WOW64_64KEY"):
            access_flags |= winreg.KEY_WOW64_64KEY

        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
            0,
            access_flags,
        )
        value, _ = winreg.QueryValueEx(key, "MachineGuid")
        return str(value).strip().lower()
    except Exception:
        return None


def build_machine_id() -> Optional[str]:
    machine_guid = get_windows_machine_guid()
    if not machine_guid:
        return None

    return hashlib.sha256(machine_guid.encode("utf-8")).hexdigest()


def get_current_machine_id() -> Optional[str]:
    return build_machine_id()


def build_activation_request() -> dict[str, Any]:
    base_dir = get_base_dir()
    machine_id = get_current_machine_id()

    return {
        "app_id": "colaboradores",
        "app_name": "Separador de Listas e Provas",
        "machine_id": machine_id,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "expected_license_path": str(get_license_path(base_dir)),
    }
    

def save_activation_request(output_path: Path) -> Path:
    request_data = build_activation_request()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(request_data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path


def validate_machine_binding(license_data: dict[str, Any]) -> bool:
    expected_machine_id = license_data.get("machine_id")
    if not expected_machine_id:
        return False

    current_machine_id = build_machine_id()
    if not current_machine_id:
        raise LicenseValidationError(
            "Não foi possível obter o identificador desta máquina."
        )

    if current_machine_id != expected_machine_id:
        raise LicenseValidationError(
            "Esta licença está vinculada a outro computador."
        )

    return True


def read_license_state(state_path: Path) -> dict[str, Any]:
    if not state_path.exists():
        return {}

    try:
        return json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def validate_clock_rollback(state_path: Path) -> datetime:
    now = datetime.now()
    state = read_license_state(state_path)

    last_run_at = state.get("last_run_at")
    if not last_run_at:
        return now

    try:
        last_run_dt = datetime.fromisoformat(last_run_at)
    except ValueError:
        return now

    tolerance = timedelta(minutes=CLOCK_ROLLBACK_TOLERANCE_MINUTES)
    if now + tolerance < last_run_dt:
        raise LicenseValidationError(
            "Foi detectado retrocesso na data/hora do Windows."
        )

    return now


def update_license_state(state_path: Path, license_data: dict[str, Any], now: datetime) -> None:
    state = {
        "last_run_at": now.isoformat(timespec="seconds"),
        "license_id": license_data.get("license_id"),
        "customer_name": license_data.get("customer_name"),
    }
    state_path.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def validate_license() -> LicenseValidationResult:
    base_dir = get_base_dir()
    license_path = get_license_path(base_dir)
    public_key_path = get_public_key_path(base_dir)
    state_path = get_state_path(base_dir)

    license_data = read_json_file(license_path)
    validate_required_fields(license_data)
    verify_signature(license_data, public_key_path)
    validate_expiration(license_data)
    machine_bound = validate_machine_binding(license_data)
    now = validate_clock_rollback(state_path)
    update_license_state(state_path, license_data, now)

    return LicenseValidationResult(
        license_id=license_data["license_id"],
        customer_name=license_data["customer_name"],
        customer_type=license_data["customer_type"],
        license_type=license_data["license_type"],
        expires_at=license_data.get("expires_at"),
        machine_bound=machine_bound,
        license_path=str(license_path),
    )


if __name__ == "__main__":
    try:
        result = validate_license()
        print("Licença válida.")
        print(f"License ID: {result.license_id}")
        print(f"Cliente: {result.customer_name}")
        print(f"Escopo: {result.customer_type}")
        print(f"Tipo: {result.license_type}")
        print(f"Expira em: {result.expires_at}")
        print(f"Vinculada à máquina: {result.machine_bound}")
        print(f"Arquivo: {result.license_path}")
    except LicenseValidationError as exc:
        print(f"ERRO DE LICENÇA: {exc}")
        raise SystemExit(1)