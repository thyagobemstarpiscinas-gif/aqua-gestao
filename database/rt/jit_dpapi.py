"""Leitura DPAPI local; nunca imprime nem persiste o texto descriptografado."""
import ctypes
from ctypes import wintypes as w
import os
from pathlib import Path
import re
import stat


class TokenRefused(RuntimeError):
    """Somente mensagem fixa, sem caminhos ou material de credencial."""


class Blob(ctypes.Structure):
    _fields_ = [("size", w.DWORD), ("data", ctypes.POINTER(ctypes.c_ubyte))]


def _check(condition):
    if not condition:
        raise TokenRefused("arquivo DPAPI privado recusado")


def _windows_acl(path):
    """Owner atual; grants somente ao owner, SYSTEM ou Administradores locais."""
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    advapi = ctypes.WinDLL("advapi32", use_last_error=True)
    ptr = ctypes.c_void_p
    kernel.GetCurrentProcess.restype = w.HANDLE
    kernel.CloseHandle.argtypes = [w.HANDLE]
    kernel.LocalFree.argtypes = [ptr]
    kernel.LocalFree.restype = ptr
    advapi.OpenProcessToken.argtypes = [w.HANDLE, w.DWORD, ctypes.POINTER(w.HANDLE)]
    advapi.OpenProcessToken.restype = w.BOOL
    advapi.GetTokenInformation.argtypes = [w.HANDLE, ctypes.c_int, ptr, w.DWORD, ctypes.POINTER(w.DWORD)]
    advapi.GetTokenInformation.restype = w.BOOL
    advapi.GetNamedSecurityInfoW.argtypes = [w.LPWSTR, ctypes.c_int, w.DWORD,
        ctypes.POINTER(ptr), ptr, ctypes.POINTER(ptr), ptr, ctypes.POINTER(ptr)]
    advapi.GetNamedSecurityInfoW.restype = w.DWORD
    advapi.EqualSid.argtypes = [ptr, ptr]
    advapi.EqualSid.restype = w.BOOL
    advapi.ConvertSidToStringSidW.argtypes = [ptr, ctypes.POINTER(ptr)]
    advapi.ConvertSidToStringSidW.restype = w.BOOL
    advapi.GetAce.argtypes = [ptr, w.DWORD, ctypes.POINTER(ptr)]
    advapi.GetAce.restype = w.BOOL
    token = w.HANDLE()
    descriptor = ptr()
    try:
        _check(advapi.OpenProcessToken(kernel.GetCurrentProcess(), 8, ctypes.byref(token)))
        needed = w.DWORD()
        advapi.GetTokenInformation(token, 1, None, 0, ctypes.byref(needed))
        _check(needed.value > 0)
        user = ctypes.create_string_buffer(needed.value)
        _check(advapi.GetTokenInformation(token, 1, user, len(user), ctypes.byref(needed)))
        user_sid = ptr.from_buffer(user).value  # TOKEN_USER starts with SID_AND_ATTRIBUTES.Sid.
        owner, dacl = ptr(), ptr()
        _check(advapi.GetNamedSecurityInfoW(str(path), 1, 5, ctypes.byref(owner), None,
            ctypes.byref(dacl), None, ctypes.byref(descriptor)) == 0)
        _check(owner.value and dacl.value and advapi.EqualSid(owner, user_sid))
        # ACL header: BYTE revision, BYTE sbz1, WORD size, WORD ace_count, WORD sbz2.
        count = ctypes.c_ushort.from_address(dacl.value + 4).value
        for index in range(count):
            ace = ptr()
            _check(advapi.GetAce(dacl, index, ctypes.byref(ace)))
            kind = ctypes.c_ubyte.from_address(ace.value).value
            flags = ctypes.c_ubyte.from_address(ace.value + 1).value
            if flags & 8 or kind == 1:  # INHERIT_ONLY or ACCESS_DENIED: no grant to this file.
                continue
            _check(kind == 0)  # Fail closed for unsupported object/callback ACEs.
            sid = ace.value + 8  # ACCESS_ALLOWED_ACE: header, mask, SidStart.
            if advapi.EqualSid(sid, user_sid):
                continue
            text = ptr()
            try:
                _check(advapi.ConvertSidToStringSidW(sid, ctypes.byref(text)))
                _check(ctypes.wstring_at(text.value) in ("S-1-5-18", "S-1-5-32-544"))
            finally:
                if text.value:
                    kernel.LocalFree(text)
    finally:
        if descriptor.value:
            kernel.LocalFree(descriptor)
        if token.value:
            kernel.CloseHandle(token)


def private_token_file(path):
    _check(os.name == "nt" and isinstance(path, str))
    p = Path(path)
    _check(p.is_absolute() and p.is_file())
    _check(not any((parent / ".git").exists() for parent in p.parents))
    for part in (p, *p.parents):
        info = part.lstat()
        _check(not part.is_symlink() and not (info.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT))
    _check(0 < p.stat().st_size <= 65536)
    _windows_acl(p)
    return p


def _ciphertext(raw):
    if raw.startswith(b"\x01\x00\x00\x00"):
        return raw, False
    text = raw.decode("utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig").strip()
    _check(bool(re.fullmatch(r"01000000[0-9a-fA-F]+", text)) and len(text) % 2 == 0)
    return bytes.fromhex(text), True


def _unprotect(raw, secure_string):
    crypto = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    crypto.CryptUnprotectData.argtypes = [ctypes.POINTER(Blob), ctypes.c_void_p,
        ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, w.DWORD, ctypes.POINTER(Blob)]
    crypto.CryptUnprotectData.restype = w.BOOL
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    buffer = ctypes.create_string_buffer(raw)
    source = Blob(len(raw), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    output = Blob()
    try:
        _check(crypto.CryptUnprotectData(ctypes.byref(source), None, None, None, None,
                                       1, ctypes.byref(output)))  # UI_FORBIDDEN; same Windows user.
        _check(output.size <= 16384)
        # Python/libpq retain immutable copies until released; no forensic zeroization claim.
        clear = ctypes.string_at(output.data, output.size)
        encoding = "utf-16-le" if secure_string or b"\x00" in clear else "utf-8"
        try:
            return clear.decode(encoding)
        finally:
            del clear
    finally:
        if output.data:
            ctypes.memset(output.data, 0, output.size)
            kernel.LocalFree(output.data)


def read_token(path):
    try:
        p = private_token_file(path)
        with p.open("rb") as stream:
            raw = stream.read(65537)
        _check(0 < len(raw) <= 65536)
        cipher, secure_string = _ciphertext(raw)
        token = _unprotect(cipher, secure_string)
        # PAT only. No JWT, plaintext file, arbitrary password, trimming or interactive fallback.
        _check(bool(re.fullmatch(r"sbp_[A-Za-z0-9_-]{20,4096}", token)))
        return token
    except Exception:
        raise TokenRefused("token DPAPI ausente, invalido ou nao privado") from None
