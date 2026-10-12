from __future__ import annotations

import base64
import hashlib
import json
import os
import platform
import re
import uuid
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Any

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import encode_dss_signature

APP_ID = "wps-sheet-automation-v1"
PRODUCT_NAME = "WPS 表格多源汇总处理器"
TRIAL_DAYS = 7
SUPPORT_EMAIL = "cmccmail@qq.com"
PUBLIC_KEY_B64 = "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEw-jKQoY80iHrnVc1DCc6DKSqaxBLNg1nPnlHaaXvRnIjj-D2B-GJ9dcK8tzUAz428MdqWu5dTD_5y2pOIJUnuQ"
_MACHINE_RE = re.compile(r"^WPSA-(?:[0-9A-F]{8}-){4}[0-9A-F]{8}$")


class LicensingError(ValueError):
    """A safe-to-display licensing or license-format error."""


@dataclass(frozen=True)
class AuthorizationStatus:
    state: str
    machine_code: str
    trial_remaining_days: int = 0
    trial_expiry: date | None = None
    license_expiry: date | None = None
    detail: str = ""


def _application_data_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "WpsSheetAutomation" / "licensing"


def trial_state_path() -> Path:
    return _application_data_dir() / "trial_state.dat"


def installed_license_path() -> Path:
    return _application_data_dir() / "license.key"


def _machine_source() -> str:
    if os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
                value, _kind = winreg.QueryValueEx(key, "MachineGuid")
            if str(value).strip():
                return "windows-machine-guid:" + str(value).strip().lower()
        except Exception:
            pass
    # The fallback avoids relying on a single network interface address.
    return "fallback:" + "|".join((
        platform.system().lower(), platform.machine().lower(), platform.node().strip().lower(),
        f"{uuid.getnode():012x}",
    ))


def machine_code(source: str | None = None) -> str:
    """Return a stable, non-reversible machine fingerprint suitable for support requests."""
    digest = hashlib.sha256((source if source is not None else _machine_source()).encode("utf-8")).hexdigest()[:40].upper()
    return "WPSA-" + "-".join(digest[index:index + 8] for index in range(0, 40, 8))


def normalize_machine_code(value: str) -> str:
    text = str(value or "").strip().upper().replace(" ", "")
    if re.fullmatch(r"[0-9A-F]{40}", text):
        text = "WPSA-" + "-".join(text[index:index + 8] for index in range(0, 40, 8))
    elif re.fullmatch(r"WPSA[0-9A-F]{40}", text):
        compact = text[4:]
        text = "WPSA-" + "-".join(compact[index:index + 8] for index in range(0, 40, 8))
    if not _MACHINE_RE.fullmatch(text):
        raise LicensingError("机器码格式无效，请从主程序授权窗口完整复制机器码。")
    return text


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _b64url_decode(text: str) -> bytes:
    if not text or not re.fullmatch(r"[A-Za-z0-9_-]+", text):
        raise LicensingError("许可证编码格式无效。")
    try:
        return base64.urlsafe_b64decode(text + "=" * (-len(text) % 4))
    except Exception as exc:
        raise LicensingError("许可证编码格式无效。") from exc


def canonical_payload(machine: str, expires: date) -> bytes:
    claims = {"app": APP_ID, "expires": expires.isoformat(), "machine": normalize_machine_code(machine)}
    return json.dumps(claims, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def generate_license(machine: str, expires: date | str, private_key_pem: bytes | str) -> str:
    """Sign a machine-bound expiry claim with a PKCS#8 P-256 private key."""
    machine = normalize_machine_code(machine)
    if isinstance(expires, str):
        try:
            expires = date.fromisoformat(expires.strip())
        except ValueError as exc:
            raise LicensingError("到期日期须为有效的 YYYY-MM-DD 日期。") from exc
    if not isinstance(expires, date):
        raise LicensingError("请选择有效的到期日期。")
    if isinstance(private_key_pem, str):
        private_key_pem = private_key_pem.encode("utf-8")
    try:
        key = serialization.load_pem_private_key(private_key_pem, password=None)
    except (ValueError, TypeError) as exc:
        raise LicensingError("无法读取私钥；请选用本程序配套的未加密 PKCS#8 PEM 私钥文件。") from exc
    if not isinstance(key, ec.EllipticCurvePrivateKey) or not isinstance(key.curve, ec.SECP256R1):
        raise LicensingError("签发私钥与本程序不匹配；需要 ECDSA P-256 私钥。")
    payload = canonical_payload(machine, expires)
    der_signature = key.sign(payload, ec.ECDSA(hashes.SHA256()))
    # WebCrypto represents ECDSA signatures as fixed-width r || s (64 bytes).
    from cryptography.hazmat.primitives.asymmetric.utils import decode_dss_signature
    r, s = decode_dss_signature(der_signature)
    raw_signature = r.to_bytes(32, "big") + s.to_bytes(32, "big")
    return _b64url_encode(payload) + "." + _b64url_encode(raw_signature)


def verify_license(token: str, machine: str, public_key_b64: str = PUBLIC_KEY_B64,
                   today: date | None = None) -> date:
    """Verify signature, product, exact machine binding, and inclusive expiry date."""
    pieces = str(token or "").strip().split(".")
    if len(pieces) != 2:
        raise LicensingError("秘钥格式不正确。")
    payload = _b64url_decode(pieces[0])
    signature = _b64url_decode(pieces[1])
    if len(signature) != 64:
        raise LicensingError("秘钥签名长度无效。")
    try:
        public_der = _b64url_decode(public_key_b64)
        public_key = serialization.load_der_public_key(public_der)
        if not isinstance(public_key, ec.EllipticCurvePublicKey) or not isinstance(public_key.curve, ec.SECP256R1):
            raise LicensingError("程序内置公钥格式错误。")
        r = int.from_bytes(signature[:32], "big")
        s = int.from_bytes(signature[32:], "big")
        public_key.verify(encode_dss_signature(r, s), payload, ec.ECDSA(hashes.SHA256()))
    except InvalidSignature as exc:
        raise LicensingError("秘钥签名无效或内容已被修改。") from exc
    except LicensingError:
        raise
    except Exception as exc:
        raise LicensingError("无法验证秘钥。") from exc
    try:
        claims: Any = json.loads(payload.decode("utf-8"))
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise LicensingError("秘钥内容格式无效。") from exc
    if not isinstance(claims, dict) or set(claims) != {"app", "expires", "machine"}:
        raise LicensingError("秘钥授权字段不完整。")
    if claims.get("app") != APP_ID:
        raise LicensingError("此秘钥不适用于本程序。")
    try:
        claim_machine = normalize_machine_code(claims.get("machine", ""))
        expiry = date.fromisoformat(claims.get("expires", ""))
    except (TypeError, ValueError, LicensingError) as exc:
        raise LicensingError("秘钥中的机器码或到期日期无效。") from exc
    if claim_machine != normalize_machine_code(machine):
        raise LicensingError("此秘钥绑定了另一台电脑的机器码。")
    if (today or date.today()) > expiry:
        raise LicensingError(f"授权已于 {expiry.isoformat()} 到期。")
    return expiry


def _protect_state(raw: bytes) -> bytes:
    if os.name != "nt":
        return b"PLAIN:" + raw
    try:
        import win32crypt
        return b"DPAPI:" + win32crypt.CryptProtectData(raw, "WpsSheetAutomation trial state", None, None, None, 0)
    except Exception as exc:
        raise LicensingError("无法保护本机试用状态；请检查 Windows 用户配置目录和系统加密服务。") from exc


def _unprotect_state(blob: bytes) -> bytes:
    if blob.startswith(b"PLAIN:") and os.name != "nt":
        return blob[6:]
    if blob.startswith(b"DPAPI:") and os.name == "nt":
        try:
            import win32crypt
            return win32crypt.CryptUnprotectData(blob[6:], None, None, None, 0)[1]
        except Exception as exc:
            raise LicensingError("本机试用记录无法解密，可能已损坏或 Windows 用户配置已更改。") from exc
    raise LicensingError("本机试用记录格式无效。")


def _load_trial_state(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    try:
        payload = json.loads(_unprotect_state(path.read_bytes()).decode("utf-8"))
    except LicensingError:
        raise
    except Exception as exc:
        raise LicensingError("本机试用记录损坏；请联系软件提供方处理。") from exc
    if not isinstance(payload, dict) or set(payload) != {"first_run", "last_seen", "machine", "version"} or payload.get("version") != 1:
        raise LicensingError("本机试用记录格式无效，不能继续试用。")
    return payload


def _save_trial_state(path: Path, state: dict[str, Any]) -> None:
    raw = json.dumps(state, sort_keys=True, separators=(",", ":")).encode("utf-8")
    protected = _protect_state(raw)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_name(path.name + ".tmp")
        temp.write_bytes(protected)
        os.replace(temp, path)
    except OSError as exc:
        raise LicensingError("无法保存本机试用记录；请检查应用数据目录的写入权限。") from exc


def evaluate_authorization(*, machine: str | None = None, today: date | None = None,
                           state_path: Path | str | None = None,
                           license_path: Path | str | None = None,
                           public_key_b64: str = PUBLIC_KEY_B64) -> AuthorizationStatus:
    machine = normalize_machine_code(machine or machine_code())
    current = today or date.today()
    state_file = Path(state_path) if state_path is not None else trial_state_path()
    key_file = Path(license_path) if license_path is not None else installed_license_path()

    try:
        saved = _load_trial_state(state_file)
        if saved is None:
            trial_state = {"version": 1, "machine": machine, "first_run": current.isoformat(), "last_seen": current.isoformat()}
        else:
            if saved.get("machine") != machine:
                return AuthorizationStatus("state_error", machine, detail="试用状态与当前机器码不匹配。")
            try:
                first_run = date.fromisoformat(saved["first_run"])
                last_seen = date.fromisoformat(saved["last_seen"])
            except (KeyError, TypeError, ValueError):
                return AuthorizationStatus("state_error", machine, detail="本机试用日期记录无效。")
            if first_run > current or last_seen > current:
                return AuthorizationStatus("clock_rollback", machine, detail=f"系统日期早于上次运行日期 {last_seen.isoformat()}；请先校正 Windows 日期和时间。")
            trial_state = dict(saved)
            trial_state["last_seen"] = current.isoformat()
        _save_trial_state(state_file, trial_state)
    except LicensingError as exc:
        return AuthorizationStatus("state_error", machine, detail=str(exc))

    license_expiry = None
    if key_file.exists():
        try:
            license_expiry = verify_license(key_file.read_text(encoding="utf-8"), machine, public_key_b64, current)
            return AuthorizationStatus("licensed", machine, license_expiry=license_expiry)
        except (OSError, UnicodeError, LicensingError):
            pass

    first_run = date.fromisoformat(trial_state["first_run"])
    trial_expiry = first_run + timedelta(days=TRIAL_DAYS)
    remaining = max(0, (trial_expiry - current).days)
    if current < trial_expiry:
        return AuthorizationStatus("trial", machine, remaining, trial_expiry, detail="试用期有效。")
    return AuthorizationStatus("expired", machine, 0, trial_expiry, detail="7 天试用期已结束，请发送机器码申请授权。")


def install_license(token: str, *, machine: str, license_path: Path | str | None = None,
                    today: date | None = None, public_key_b64: str = PUBLIC_KEY_B64) -> date:
    expiry = verify_license(token, machine, public_key_b64, today)
    path = Path(license_path) if license_path is not None else installed_license_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_name(path.name + ".tmp")
        temp.write_text(token.strip(), encoding="utf-8")
        os.replace(temp, path)
    except OSError as exc:
        raise LicensingError("无法保存授权文件；请检查用户配置目录权限。") from exc
    return expiry


def authorize_startup(root: Any) -> bool:
    """Display the trial/activation dialog; return True only when app may open."""
    import tkinter as tk
    from tkinter import messagebox

    status = evaluate_authorization()
    if status.state == "licensed":
        return True

    dialog = tk.Toplevel(root)
    dialog.title("软件授权")
    dialog.resizable(False, False)
    # The main root is intentionally withdrawn during startup. On Windows, a
    # transient dialog owned by a hidden root can itself remain invisible.
    # Only attach to the root when it is actually visible.
    try:
        if root.state() != "withdrawn":
            dialog.transient(root)
    except Exception:
        pass
    result = {"authorized": False}
    machine_var = tk.StringVar(value=status.machine_code)
    token_var = tk.StringVar()

    if status.state == "trial":
        headline = f"当前为 7 天试用期，剩余 {status.trial_remaining_days} 天；试用截止：{status.trial_expiry.isoformat()}。"
    elif status.state == "clock_rollback":
        headline = status.detail + "\n校正系统日期后重新启动程序。"
    elif status.state == "state_error":
        headline = status.detail + "\n请联系软件提供方处理试用记录。"
    else:
        headline = "7 天试用期已到期，请复制机器码并发送至邮箱申请秘钥。"

    tk.Label(dialog, text=PRODUCT_NAME, font=("Microsoft YaHei UI", 12, "bold")).pack(padx=18, pady=(16, 8), anchor="w")
    tk.Label(dialog, text=headline, justify="left", wraplength=570).pack(padx=18, pady=(0, 10), anchor="w")
    tk.Label(dialog, text="本机机器码：").pack(padx=18, anchor="w")
    machine_entry = tk.Entry(dialog, textvariable=machine_var, width=60, state="readonly", readonlybackground="white")
    machine_entry.pack(padx=18, pady=(3, 8), fill="x")

    def copy_machine() -> None:
        dialog.clipboard_clear()
        dialog.clipboard_append(status.machine_code)
        dialog.update_idletasks()
        messagebox.showinfo("软件授权", "机器码已复制。请将机器码发送至 " + SUPPORT_EMAIL, parent=dialog)

    tk.Button(dialog, text="复制机器码", command=copy_machine).pack(padx=18, pady=(0, 10), anchor="w")
    tk.Label(dialog, text=f"申请授权：将上方机器码发送到 {SUPPORT_EMAIL}。\n收到秘钥后粘贴到下方：", justify="left").pack(padx=18, anchor="w")
    token_entry = tk.Entry(dialog, textvariable=token_var, width=80)
    token_entry.pack(padx=18, pady=(5, 10), fill="x")

    can_continue = status.state == "trial"

    def activate() -> None:
        try:
            expiry = install_license(token_var.get(), machine=status.machine_code)
            result["authorized"] = True
            messagebox.showinfo("软件授权", f"授权成功，到期日期：{expiry.isoformat()}。", parent=dialog)
            dialog.destroy()
        except LicensingError as exc:
            messagebox.showerror("软件授权", str(exc), parent=dialog)

    def enter_trial() -> None:
        result["authorized"] = True
        dialog.destroy()

    buttons = tk.Frame(dialog)
    buttons.pack(pady=(3, 16))
    if can_continue:
        tk.Button(buttons, text="继续试用", command=enter_trial, width=14).pack(side="left", padx=5)
    if status.state not in {"clock_rollback", "state_error"}:
        tk.Button(buttons, text="激活秘钥", command=activate, width=14).pack(side="left", padx=5)
    tk.Button(buttons, text="退出", command=dialog.destroy, width=12).pack(side="left", padx=5)
    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    dialog.update_idletasks()
    x = max(0, (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2)
    y = max(0, (dialog.winfo_screenheight() - dialog.winfo_height()) // 2)
    dialog.geometry(f"+{x}+{y}")
    dialog.deiconify()
    dialog.wait_visibility()
    dialog.lift()
    dialog.grab_set()
    token_entry.focus_set()
    root.wait_window(dialog)
    return result["authorized"]
