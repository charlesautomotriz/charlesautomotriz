# App Streamlit - Cotizaciones Charles Servicio Automotriz

## Archivos incluidos
- `app.py`: aplicación principal en Streamlit.
- `Cotizacion Charles Servicio Automotriz.docx`: plantilla Word usada para generar la cotización.
- `marca charles.png`: logo superior.
- `logo charles blanco 21 jun 2025, 23_06_15.png`: marca de agua visual de la app.
- `requirements.txt`: librerías Python.
- `packages.txt`: instala LibreOffice en Streamlit Cloud para convertir Word a PDF.

## Cómo ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Qué guarda la aplicación
- Base SQLite: `cotizaciones_charles.db`.
- Base Excel sincronizada: `base_cotizaciones_charles.xlsx`.
- Documentos generados: carpeta `cotizaciones_generadas`.

## Notas importantes
- La plantilla entregada usa textos tipo `<<Nombre del cliente>>`; el código los reemplaza directamente en el Word.
- La plantilla trae espacio visible para 3 trabajos y 3 repuestos. Si agregas más ítems, se incorporan en Observaciones.
- El total final suma: mano de obra + repuestos + otros.
- En Streamlit Cloud, el archivo Excel/SQLite puede reiniciarse si la app se redepliega. Para una base permanente, conviene conectar Google Sheets o una base externa.
