import os
from pyhanko.sign import signers
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign.fields import SigSeedSubFilter

def firmar_pdf(ruta_pdf_entrada, ruta_pfx, password_pfx, motivo="Certificación Académica", ubicacion="Burgos, España"):
    """
    Firma digitalmente un PDF usando un certificado .pfx/.p12.
    Versión robustecida usando el cargador oficial de pyHanko.
    """
    if not os.path.exists(ruta_pdf_entrada):
        print(f"❌ Error: No existe el PDF {ruta_pdf_entrada}")
        return False

    if not os.path.exists(ruta_pfx):
        print(f"❌ Error: No existe el certificado {ruta_pfx}")
        return False

    try:
        # 1. Cargar el firmante usando el método de alto nivel de pyHanko
        # Esto maneja la decodificación del pfx de forma más segura.
        signer = signers.SimpleSigner.load_pkcs12(
            pfx_file=ruta_pfx,
            passphrase=password_pfx.encode('utf-8')
        )

        # 2. Preparar el archivo PDF
        with open(ruta_pdf_entrada, 'rb') as inf:
            w = IncrementalPdfFileWriter(inf)
            
            # 3. Metadatos de la firma (PAdES es el estándar europeo)
            meta = signers.PdfSignatureMetadata(
                field_name='FirmaDigitalUBU',
                reason=motivo,
                location=ubicacion,
                subfilter=SigSeedSubFilter.PADES
            )

            # 4. Proceso de firma (sobrescribiendo el archivo de forma segura)
            ruta_temp = ruta_pdf_entrada + ".signed.tmp"
            
            with open(ruta_temp, 'wb') as outf:
                signers.sign_pdf(
                    w, meta, signer=signer, output=outf,
                )
        
        # 5. Si todo fue bien, reemplazar el original
        os.replace(ruta_temp, ruta_pdf_entrada)
        return True

    except Exception as e:
        # Imprimir el error real en la consola para depurar
        print(f"!!!!!!!!!! ERROR REAL DE FIRMA: {e} !!!!!!!!!!")
        # Limpiar archivo temporal si falló
        if os.path.exists(ruta_pdf_entrada + ".signed.tmp"):
            os.remove(ruta_pdf_entrada + ".signed.tmp")
        return False