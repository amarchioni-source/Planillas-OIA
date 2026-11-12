# Generador de Planillas Angus

App web para procesar archivos ROMANEO ("Detalle de cargas remito ...") y generar planillas completadas con cortes AA.

## Estructura

```
planillas-app/
├── app.py                  # Servidor Flask
├── Planilla_BASE_.xlsx     # Template base (se usa siempre esta)
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

- Subís uno o varios archivos ROMANEO (.xlsx)
- Subís el Packing List (PDF) de cada romaneo: de ahí sale el CONSIGNEE (celda E25)
- Procesás y descargás cada planilla o todas juntas en ZIP

## Reglas de procesamiento

- Columnas del romaneo usadas: `Producto`, `Product`, `Cod Prod`, `Cartons`, `Net W`, `Gross`, `Tropa`, `Prod Date`
- Se agrupa por `Producto`: cajas = suma `Cartons`, neto = suma `Net W`, bruto = suma `Gross`
- Descripción en inglés (columna E) = columna `Product` del romaneo
- Lotes (E70) = fechas de `Prod Date`
- Consignee (E25) = bloque CONSIGNEE del Packing List. Cada romaneo se empareja con su Packing List por códigos de producto (CDxxxxx) y total de cajas; si se sube un solo romaneo y un solo Packing List, se usan juntos directamente

- Nombre del archivo: `PLANILLA OIA <N° de venta>.xlsx`, con el N° de venta del Packing List (EXPORTER REFERENCE). Sin Packing List: `PLANILLA OIA <nombre del romaneo>.xlsx`
- Solo incluye cortes cuyo Producto contenga **AA** (Angus)
- Cortes se escriben en filas **34 a 63** (máx 30 cortes)
- Tropas en **E69**, lotes en **E70**
