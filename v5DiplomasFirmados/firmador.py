import os
from pyhanko.sign import signers, fields
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign.fields import SigSeedSubFilter

def firmar_pdf(ruta_pdf_entrada, ruta_pfx, password_pfx, motivo="Certificación Académica", ubicacion="Burgos, España"):
    """
    Firma digitalmente un PDF usando un certificado .pfx/.p12
    La firma es INVISIBLE (criptográfica) pero valida el documento en Adobe.
    """
    if not os.path.exists(ruta_pdf_entrada):
        print(f"❌ Error: No existe el PDF {ruta_pdf_entrada}")
        return False

    if not os.path.exists(ruta_pfx):
        print(f"❌ Error: No existe el certificado {ruta_pfx}")
        return False

    try:
        # 1. Cargar el firmante desde el archivo PFX
        # pyHanko maneja la criptografía compleja aquí
        signer = signers.P12Signer(
            load_p12(ruta_pfx, password_pfx.encode('utf-8'))
        )

        # 2. Preparar el archivo para escritura incremental (más seguro)
        with open(ruta_pdf_entrada, 'rb') as inf:
            w = IncrementalPdfFileWriter(inf)
            
            # 3. Configurar metadatos de la firma
            meta = signers.PdfSignatureMetadata(
                field_name='FirmaDigitalUBU',
                reason=motivo,
                location=ubicacion,
                subfilter=SigSeedSubFilter.PADES  # Estándar PAdES (Long Term Validation)
            )

            # 4. Sobrescribir el archivo original con la versión firmada
            # (Usamos un archivo temporal y luego renombramos para evitar corrupciones)
            ruta_temp = ruta_pdf_entrada + ".signed"
            
            with open(ruta_temp, 'wb') as outf:
                signers.sign_pdf(
                    w, meta, signer=signer, output=outf,
                )
        
        # Reemplazar el original por el firmado
        os.replace(ruta_temp, ruta_pdf_entrada)
        return True

    except Exception as e:
        print(f"❌ Error al firmar {os.path.basename(ruta_pdf_entrada)}: {e}")
        # Limpieza si falló
        if os.path.exists(ruta_pdf_entrada + ".signed"):
            os.remove(ruta_pdf_entrada + ".signed")
        return False

# Función auxiliar para cargar p12 de forma segura
def load_p12(fname, password):
    from pyhanko.sign.pkcs11 import PKCS11Signer
    from cryptography.hazmat.primitives.serialization import pkcs12
    
    with open(fname, 'rb') as f:
        p12_data = f.read()
    
    # Cargar clave privada y certificado
    private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
        p12_data, password
    )
    return signers.SimpleSigner(
        signing_cert=certificate,
        signing_key=private_key,
        cert_registry=signers.SimpleCertificateStore.from_certs(
            [certificate] + (additional_certificates or [])
        )
    )