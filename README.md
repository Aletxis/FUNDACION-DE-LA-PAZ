# Observatorio - Movilidad Humana (Dashboard)

Dashboard exploratorio construido con Streamlit para el Reporte Consolidado de Movilidad Humana.

## ⚠️ Datos

Este repositorio **no incluye ni descarga automáticamente** ningún archivo de datos. El dashboard exige subir `REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx` manualmente en cada sesión, vía un `st.file_uploader` dentro de la propia app.

Esto es intencional: el archivo contiene información personal identificable (PII) de personas en situación de movilidad humana, incluyendo datos de niños, niñas y adolescentes (NNA) y sus tutores. No puede vivir en Git, ni en un link de Drive "compartido", ni en ningún storage externo con acceso automático. El archivo subido vive solo en la memoria (RAM) de esa sesión del navegador — nunca se guarda en el servidor ni se sube a ningún otro lugar.

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

Al abrir el dashboard, pedirá subir el `.xlsx` antes de mostrar cualquier gráfico.

## Estructura

- `app.py` — lógica del dashboard (carga vía upload, cacheo en memoria con `st.cache_data`, visualizaciones).
- `logo1.jpeg`, `logo2.jpeg` — logos institucionales.
- `REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx` — **nunca incluido en el repo**, se sube manualmente en cada sesión (ver arriba).

## Despliegue en Streamlit Cloud

1. Conecta este repo (recomendado: **privado**) en share.streamlit.io.
2. En Settings → Sharing, restringe el acceso a personas específicas (tutora/revisor).
3. No hace falta configurar ningún secreto: cada persona que abra el link sube su propia copia del `.xlsx` al entrar.
