# -*- coding: utf-8 -*-
"""
License check - activation and validity verification.
"""
import os
import sys
import json
import base64
import hashlib
import datetime

AUTHOR_PHONE = ""
AUTHOR_NAME = "Administrator"
LICENSE_DAYS = 365  # 1 year

def _get_app_dir():
    """Get the directory where the exe/script is located."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _license_file():
    return os.path.join(_get_app_dir(), '.license')

def _encode(data):
    """Simple obfuscation for license data."""
    raw = json.dumps(data, ensure_ascii=False)
    sig = hashlib.md5((raw + "DebtMailPro2026").encode()).hexdigest()[:8]
    payload = raw + "|" + sig
    return base64.b64encode(payload.encode()).decode()

def _decode(encoded):
    """Decode obfuscated license data."""
    try:
        payload = base64.b64decode(encoded.encode()).decode()
        raw, sig = payload.rsplit("|", 1)
        expected = hashlib.md5((raw + "DebtMailPro2026").encode()).hexdigest()[:8]
        if sig != expected:
            return None
        return json.loads(raw)
    except Exception:
        return None

def activate():
    """Create activation record on first run."""
    lf = _license_file()
    if os.path.exists(lf):
        return  # Already activated
    data = {
        'activated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'days': LICENSE_DAYS
    }
    with open(lf, 'w', encoding='utf-8') as f:
        f.write(_encode(data))

def check_license():
    """
    Check license validity.
    Returns: (is_valid: bool, message: str, days_remaining: int)
    """
    lf = _license_file()
    
    # First run - activate
    if not os.path.exists(lf):
        activate()
        return True, "Kích hoạt thành công!", LICENSE_DAYS
    
    # Read license
    try:
        with open(lf, 'r', encoding='utf-8') as f:
            encoded = f.read().strip()
    except Exception:
        activate()
        return True, "Kích hoạt lại thành công!", LICENSE_DAYS
    
    data = _decode(encoded)
    if data is None:
        activate()
        return True, "License reset.", LICENSE_DAYS
    
    activated = datetime.datetime.strptime(data['activated'], '%Y-%m-%d %H:%M:%S')
    expiry = activated + datetime.timedelta(days=data.get('days', LICENSE_DAYS))
    now = datetime.datetime.now()
    remaining = (expiry - now).days
    
    if remaining <= 0:
        msg = (
            f"⚠️ PHẦN MỀM ĐÃ HẾT HẠN SỬ DỤNG!\n\n"
            f"Ngày kích hoạt: {data['activated']}\n"
            f"Ngày hết hạn: {expiry.strftime('%Y-%m-%d')}\n\n"
            f"Vui lòng liên hệ quản trị viên để gia hạn.\n"
        )
        return False, msg, 0
    
    if remaining <= 30:
        msg = f"⚠️ Còn {remaining} ngày sử dụng. Vui lòng liên hệ quản trị viên để gia hạn."
    else:
        msg = f"License OK - còn {remaining} ngày."
    
    return True, msg, remaining

def get_license_info():
    """Get license info for display in UI."""
    lf = _license_file()
    if not os.path.exists(lf):
        return None
    try:
        with open(lf, 'r', encoding='utf-8') as f:
            encoded = f.read().strip()
        data = _decode(encoded)
        if data:
            activated = datetime.datetime.strptime(data['activated'], '%Y-%m-%d %H:%M:%S')
            expiry = activated + datetime.timedelta(days=data.get('days', LICENSE_DAYS))
            remaining = (expiry - datetime.datetime.now()).days
            return {
                'activated': data['activated'],
                'expiry': expiry.strftime('%Y-%m-%d'),
                'remaining': max(remaining, 0),
                'expired': remaining <= 0
            }
    except Exception:
        pass
    return None
