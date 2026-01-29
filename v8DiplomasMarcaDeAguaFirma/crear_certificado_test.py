from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.serialization import pkcs12  # <--- CORREGIDO AQUÍ
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import hashes
from cryptography.x509.oid import NameOID
from cryptography import x509
import datetime
import os

def generar_pfx_prueba():
    print("Generando certificado de prueba...")
    
    # 1. Crear clave privada
    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    # 2. Datos del certificado (Falsos para pruebas)
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, u"ES"),
        x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, u"Burgos"),
        x509.NameAttribute(NameOID.LOCALITY_NAME, u"Burgos"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, u"Universidad de Burgos (TEST)"),
        x509.NameAttribute(NameOID.COMMON_NAME, u"Pruebas Generador Diplomas"),
    ])

    # 3. Crear el certificado
    cert = x509.CertificateBuilder().subject_name(
        subject
    ).issuer_name(
        issuer
    ).public_key(
        key.public_key()
    ).serial_number(
        x509.random_serial_number()
    ).not_valid_before(
        datetime.datetime.utcnow()
    ).not_valid_after(
        # Valido por 1 año
        datetime.datetime.utcnow() + datetime.timedelta(days=365)
    ).add_extension(
        x509.BasicConstraints(ca=True, path_length=None), critical=True,
    ).sign(key, hashes.SHA256())

    # 4. Guardar como .pfx
    # Contraseña para el archivo: "1234"
    password = b"1234"
    
    # USAMOS pkcs12 DIRECTAMENTE
    pfx_data = pkcs12.serialize_key_and_certificates(
        name=b"Certificado Pruebas",
        key=key,
        cert=cert,
        cas=None,
        encryption_algorithm=serialization.BestAvailableEncryption(password)
    )

    with open("certificado_prueba.pfx", "wb") as f:
        f.write(pfx_data)
        
    print("✅ ¡Éxito! Se ha creado 'certificado_prueba.pfx'")
    print("🔑 La contraseña es: 1234")

if __name__ == "__main__":
    generar_pfx_prueba()