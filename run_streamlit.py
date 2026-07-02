import runpy
import ssl
import sys


def safe_create_default_context(
    purpose=ssl.Purpose.SERVER_AUTH,
    *,
    cafile=None,
    capath=None,
    cadata=None,
):
    protocol = (
        ssl.PROTOCOL_TLS_SERVER
        if purpose == ssl.Purpose.CLIENT_AUTH
        else ssl.PROTOCOL_TLS_CLIENT
    )
    context = ssl.SSLContext(protocol)

    if cafile or capath or cadata:
        context.load_verify_locations(cafile=cafile, capath=capath, cadata=cadata)
    else:
        context.load_default_certs(purpose)

    return context


def safe_create_unverified_context(*args, **kwargs):
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    return context


ssl.create_default_context = safe_create_default_context
ssl._create_default_https_context = safe_create_default_context
ssl._create_unverified_context = safe_create_unverified_context


def safe_create_urllib3_context(*args, **kwargs):
    cert_reqs = kwargs.get("cert_reqs", ssl.CERT_REQUIRED)
    if cert_reqs is None:
        cert_reqs = ssl.CERT_REQUIRED

    context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    if cert_reqs == ssl.CERT_NONE:
        context.check_hostname = False
        context.verify_mode = cert_reqs
    else:
        context.verify_mode = cert_reqs
        context.check_hostname = True

    return context


try:
    import urllib3.util.ssl_ as urllib3_ssl

    urllib3_ssl.create_urllib3_context = safe_create_urllib3_context
except Exception:
    pass


if __name__ == "__main__":
    streamlit_args = sys.argv[1:] or [
        "run",
        "audit_app.py",
        "--global.developmentMode",
        "false",
        "--server.headless",
        "true",
        "--server.port",
        "8501",
        "--browser.gatherUsageStats",
        "false",
    ]

    sys.argv = ["streamlit", *streamlit_args]
    runpy.run_module("streamlit", run_name="__main__")
