# App Streamlit - Cotizaciones Charles Servicio Automotriz

Aplicación para generar cotizaciones de reparación de vehículos, con historial, correlativos, estados, catálogo de servicios/repuestos y dashboard.

## Ejecutar localmente o en Codespaces

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## PDF

Para PDF se usa LibreOffice/soffice.

En Codespaces:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice
```

En Streamlit Cloud, el archivo `packages.txt` ya incluye:

```txt
libreoffice
```

## Base persistente recomendada: Google Sheets / Drive

La app funciona en modo local con SQLite, pero para producción conviene usar Google Sheets para no perder historial ni correlativos al reiniciar o redeplegar la app.

Crear un Google Sheet y compartirlo con el correo del service account. Luego agregar esto en `.streamlit/secrets.toml` o en los secrets de Streamlit Cloud:

```toml
[gsheets]
spreadsheet_id = "ID_DE_TU_GOOGLE_SHEET"

[gsheets.service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "nombre@proyecto.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

## Qué guarda la aplicación

En Google Sheets o SQLite se guardan:

- `cotizaciones`: cabecera, cliente, vehículo, totales, estado y correlativo.
- `detalle_items`: mano de obra, repuestos y otros.
- `catalogo`: listado editable para seleccionar rápido mano de obra, repuestos y otros.

La app ya no guarda documentos Word/PDF como archivos permanentes. Los genera solo para descarga inmediata, y también puede recrearlos desde el historial usando la base.
