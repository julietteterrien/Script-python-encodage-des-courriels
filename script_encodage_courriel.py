import os
import email
import chardet
from email import policy
from email.generator import BytesGenerator
from email.header import decode_header, make_header


# Chemins vers le dossier d'entrée et le dossier de sortie
input_dir = "/nom-du-dossier-contenant-les-eml"
output_dir = "./eml_utf8"
os.makedirs(output_dir, exist_ok=True)


# Détection automatique de l'encodage
def detect_encoding(data):
    return chardet.detect(data)["encoding"] or "utf-8"


# Décodage des en-têtes MIME (ex: Subject, From...)
def decode_mime_header(header_val):
    if header_val:
        try:
            return str(make_header(decode_header(header_val)))
        except Exception as e:
            print(f"[!] Erreur de décodage d'en-tête : {e}")
    return header_val


# Traitement des fichiers .eml
for filename in os.listdir(input_dir):
    if not filename.lower().endswith(".eml"):
        continue


    input_path = os.path.join(input_dir, filename)
    with open(input_path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)


    # Décodage des en-têtes
    for header in ["Subject", "From", "To", "Cc", "Bcc"]:
        if msg[header]:
            msg.replace_header(header, decode_mime_header(msg[header]))


    # Traitement du corps et des pièces jointes
    for part in msg.walk():
        # -- Texte --
        if part.get_content_maintype() == "text":
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or detect_encoding(payload)
                try:
                    decoded = payload.decode(charset)
                except Exception:
                    decoded = payload.decode(charset, errors="replace")
                part.set_payload(decoded, charset="utf-8")


        # -- Pièce jointe : nom décodé --
        attachment_name = part.get_filename()
        if attachment_name:
            try:
                decoded_filename = decode_mime_header(attachment_name)
                if decoded_filename:
                    part.set_param("filename", decoded_filename, header="Content-Disposition")
                    if "name=" in str(part.get("Content-Type", "")):
                        part.set_param("name", decoded_filename, header="Content-Type")
            except Exception as e:
                print(f"[!] Erreur de décodage du nom de fichier joint ({attachment_name}) : {e}")


    # Sauvegarde de l'email converti
    output_path = os.path.join(output_dir, filename)
    with open(output_path, "wb") as f:
        BytesGenerator(f, policy=policy.default).flatten(msg)


print("Tous les fichiers .eml ont été convertis en UTF-8 avec les en-têtes et les noms de pièces jointes décodés.")