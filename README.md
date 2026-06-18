# Generador de Planillas Angus

App web para procesar archivos PIQUEO y generar planillas completadas con cortes AA.

## Estructura

```
planillas-app/
├── app.py                  # Servidor Flask
├── Planilla_BASE_.xlsx     # Template base (subir acá)
├── requirements.txt
├── Procfile
└── templates/
    └── index.html
```

## Deploy en Render

1. **Subí esta carpeta a un repositorio GitHub** (puede ser privado)
2. Incluí el archivo `Planilla_BASE_.xlsx` en la raíz del repo
3. En [render.com](https://render.com):
   - New → Web Service
   - Conectá tu repo de GitHub
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
   - Plan: Free
4. Deploy ✓

## Uso

- Subís uno o varios archivos PIQUEO (.xlsx)
- Opcionalmente subís la Planilla_BASE_ (si no está en el servidor)
- Procesás y descargás cada planilla o todas juntas en ZIP

## Reglas de procesamiento

- Solo incluye cortes cuyo Producto contenga **AA** (Angus)
- Cortes se escriben en filas **34 a 63** (máx 30 cortes)
- Tropas en **E69**, lotes en **E70**
