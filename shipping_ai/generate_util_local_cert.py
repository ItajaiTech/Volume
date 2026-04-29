#!/usr/bin/env python3
from datetime import datetime, timedelta, UTC
from ipaddress import ip_address
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID


BASE_DIR = Path(__file__).resolve().parent
CERTS_DIR = BASE_DIR / "certs"
KEY_FILE = CERTS_DIR / "util.local.key"
CERT_FILE = CERTS_DIR / "util.local.crt"


def build_san_entries(hosts):
    entries = []
    for host in hosts:
        value = str(host).strip()
        if not value:
            continue
        try:
            entries.append(x509.IPAddress(ip_address(value)))
        except ValueError:
            entries.append(x509.DNSName(value))
    return entries


def generate_certificate(hosts=None):
    if hosts is None:
        hosts = ["util.local", "volume.local", "localhost", "127.0.0.1"]

    CERTS_DIR.mkdir(parents=True, exist_ok=True)

    private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    subject = issuer = x509.Name(
        [
            x509.NameAttribute(NameOID.COUNTRY_NAME, "BR"),
            x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, "Local"),
            x509.NameAttribute(NameOID.LOCALITY_NAME, "Local"),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Volume"),
            x509.NameAttribute(NameOID.COMMON_NAME, hosts[0]),
        ]
    )

    now = datetime.now(UTC)
    cert_builder = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(private_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + timedelta(days=825))
        .add_extension(x509.SubjectAlternativeName(build_san_entries(hosts)), critical=False)
    )

    certificate = cert_builder.sign(private_key=private_key, algorithm=hashes.SHA256())

    KEY_FILE.write_bytes(
        private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    CERT_FILE.write_bytes(certificate.public_bytes(serialization.Encoding.PEM))

    print(f"Certificado gerado: {CERT_FILE}")
    print(f"Chave privada gerada: {KEY_FILE}")
    print(f"Hosts cobertos: {', '.join(hosts)}")


if __name__ == "__main__":
    generate_certificate()