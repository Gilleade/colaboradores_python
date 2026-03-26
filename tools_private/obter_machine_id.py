import hashlib
import platform


def get_windows_machine_guid() -> str | None:
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


def build_machine_id() -> str | None:
    machine_guid = get_windows_machine_guid()
    if not machine_guid:
        return None

    return hashlib.sha256(machine_guid.encode("utf-8")).hexdigest()


def main() -> None:
    machine_id = build_machine_id()
    if not machine_id:
        print("Não foi possível obter a machine_id desta máquina.")
        raise SystemExit(1)

    print("Machine ID desta máquina:")
    print(machine_id)


if __name__ == "__main__":
    main()