# Observatorio - Movilidad Humana (Dashboard)

Dashboard exploratorio construido con Streamlit para el Reporte Consolidado de Movilidad Humana.

## ⚠️ Datos

Este repositorio **no incluye la base de datos real** (`REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx`) porque contiene información personal identificable (PII) de personas en situación de movilidad humana, incluyendo datos de niños, niñas y adolescentes (NNA) y sus tutores. Ese archivo se maneja localmente y está excluido vía `.gitignore`.

Para ejecutar el dashboard necesitas colocar tu propio archivo `REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx` (mismo esquema de columnas) en la raíz del proyecto, junto a `app.py`.

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

## Estructura

- `app.py` — lógica del dashboard (carga de datos, cacheo, visualizaciones).
- `logo1.jpeg`, `logo2.jpeg` — logos institucionales.
- `REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx` — **no incluido**, datos sensibles (ver arriba).
