"""
app.py - Generador de Planillas Angus
Flask app para procesar archivos ROMANEO y generar planillas completadas
"""
import math, os, re, shutil, zipfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = Path('/tmp/uploads')
OUTPUT_FOLDER = Path('/tmp/outputs')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

PLANILLA_BASE = Path(__file__).parent / 'Planilla_BASE_.xlsx'


def find_data_sheet_and_header(filepath, xl):
    """
    Busca la hoja Y fila de encabezado que contenga 'Producto'.
    Prueba todas las hojas en orden de prioridad.
    """
    def priority(name):
        n = name.lower()
        if 'hoja' in n:     return 0
        if 'pegar' in n:    return 1
        if 'extracto' in n: return 9
        return 5

    for sheet in sorted(xl.sheet_names, key=priority):
        raw = pd.read_excel(filepath, sheet_name=sheet, header=None)
        for i, row in raw.iterrows():
            for val in row.values:
                if str(val).strip().lower() == 'producto':
                    return sheet, i

    raise ValueError(
        f"No se encontró columna 'Producto' en ninguna hoja. "
        f"Hojas disponibles: {xl.sheet_names}"
    )


def extract_tropas(val):
    if pd.isna(val):
        return set()
    found = re.findall(r'\(\d+\)(\d+)', str(val))
    if found:
        return set(found)
    return set(re.findall(r'\d+', str(val)))


# Columnas del ROMANEO ("Detalle de cargas remito ...")
COL_PRODUCTO = 'Producto'     # descripción en castellano (se usa para agrupar y filtrar AA)
COL_PRODUCT_EN = 'Product'    # descripción en inglés que ya trae el romaneo
COL_CODIGO = 'Cod Prod'
COL_CAJAS = 'Cartons'
COL_NETO = 'Net W'
COL_BRUTO = 'Gross'
COL_TROPA = 'Tropa'
COL_FECHA_PROD = 'Prod Date'

COLUMNAS_REQUERIDAS = [COL_PRODUCTO, COL_CAJAS, COL_NETO, COL_BRUTO, COL_TROPA, COL_FECHA_PROD]


def _clean_code(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return '' if s.lower() == 'nan' else s


# ---------------------------------------------------------------------------
# PACKING LIST (PDF): de acá sale el CONSIGNEE de cada carga
# ---------------------------------------------------------------------------
# Etiquetas de la columna derecha del encabezado del packing list; marcan el
# fin del bloque CONSIGNEE.
_PL_FIN_CONSIGNEE = re.compile(
    r'^(INVOICE DATE|BUYERS REF|EXPORTER REFERENCE|ESTABLISHMENT NO|VESSEL/AIRCRAFT|'
    r'VOYAGE NO|DEPARTURE DATE|LOADING PORT|DISCHARGE PORT|SHIPPING MARKS|NOTIFY)',
    re.IGNORECASE,
)


def parse_packing_list(pdf_path):
    """
    Lee un PACKING LIST en PDF y devuelve:
      consignee: texto del bloque CONSIGNEE (una línea por renglón)
      venta:     EXPORTER REFERENCE (N° de venta)
      codigos:   códigos de producto (CDxxxxx) que aparecen
      cajas:     total de cajas (TOTAL : N CTNS)
    """
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    text = '\n'.join((p.extract_text() or '') for p in reader.pages)
    lines = [l.strip() for l in text.splitlines()]

    consignee_lines = []
    try:
        start = next(i for i, l in enumerate(lines) if l.upper() == 'CONSIGNEE')
    except StopIteration:
        start = None
    if start is not None:
        for l in lines[start + 1:]:
            if _PL_FIN_CONSIGNEE.match(l):
                break
            if l:
                consignee_lines.append(l)

    m_venta = re.search(r'EXPORTER REFERENCE\s+(\S+)', text, re.IGNORECASE)
    m_cajas = re.search(r'TOTAL\s*:\s*([\d.,]+)\s*CTNS', text, re.IGNORECASE)
    cajas = None
    if m_cajas:
        try:
            cajas = int(float(m_cajas.group(1).replace(',', '')))
        except ValueError:
            cajas = None

    return {
        'archivo': Path(pdf_path).name,
        'consignee': '\n'.join(consignee_lines),
        'venta': m_venta.group(1) if m_venta else None,
        'codigos': set(re.findall(r'\b(CD\d+)\b', text)),
        'cajas': cajas,
    }


def match_packing_list(packing_lists, codigos, cajas, solo_uno=False):
    """
    Elige el packing list que corresponde a un romaneo: mismo total de cajas
    y mayor coincidencia de códigos de producto. Si hay un solo romaneo y un
    solo packing list, se usan juntos directamente.
    """
    if not packing_lists:
        return None
    if solo_uno and len(packing_lists) == 1:
        return packing_lists[0]

    mejor, mejor_score = None, 0
    for pl in packing_lists:
        comunes = len(codigos & pl['codigos'])
        if not comunes:
            continue
        score = comunes + (1000 if pl['cajas'] == cajas else 0)
        if score > mejor_score:
            mejor, mejor_score = pl, score
    return mejor


def _nombre_salida(venta):
    """
    'PLANILLA OIA <N° de venta>.xlsx' con el N° de venta del Packing List
    (EXPORTER REFERENCE). Si el nombre ya existe en esta tanda, se agrega (2), (3)...
    """
    base = 'PLANILLA OIA ' + re.sub(r'[\\/:*?"<>|]', '-', venta)
    nombre = f'{base}.xlsx'
    n = 2
    while (OUTPUT_FOLDER / nombre).exists():
        nombre = f'{base} ({n}).xlsx'
        n += 1
    return nombre


def process_file(filepath, packing_lists=None, solo_uno=False):
    filepath = Path(filepath)

    with pd.ExcelFile(filepath) as xl:
        data_sheet, header_row = find_data_sheet_and_header(filepath, xl)

    df = pd.read_excel(filepath, sheet_name=data_sheet, header=header_row)

    # Normalizar nombres de columna
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"El archivo no parece un ROMANEO: faltan columnas {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    # Descartar filas vacías / totales (sin producto)
    df = df[df[COL_PRODUCTO].notna()]

    # Filtro Angus: solo productos con AA en columna Producto
    df = df[df[COL_PRODUCTO].astype(str).str.contains(r'\bAA\b', na=False)]

    for c in (COL_CAJAS, COL_NETO, COL_BRUTO):
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    grouped = (
        df.groupby(COL_PRODUCTO, sort=True)
        .agg(Cajas=(COL_CAJAS, 'sum'), Peso_kg=(COL_NETO, 'sum'), Peso_Bruto=(COL_BRUTO, 'sum'))
        .reset_index()
        .rename(columns={COL_PRODUCTO: 'Producto'})
    )

    # Descripción en inglés: columna 'Product' del romaneo (si falta, Producto original)
    first = df.groupby(COL_PRODUCTO).first()

    def get_description(producto):
        if COL_PRODUCT_EN in first.columns:
            en = str(first.loc[producto, COL_PRODUCT_EN]).strip()
            if en and en.lower() != 'nan':
                return en
        return producto

    grouped['Descripcion'] = grouped['Producto'].apply(get_description)

    all_tropas = set()
    for t in df[COL_TROPA].dropna():
        all_tropas.update(extract_tropas(t))
    tropas_str = ', '.join(sorted(all_tropas, key=lambda x: int(x)))

    all_lotes = set()
    for d in df[COL_FECHA_PROD].dropna():
        all_lotes.add(pd.Timestamp(d).normalize())
    lotes_str = ', '.join(d.strftime('%d/%m/%Y') for d in sorted(all_lotes))

    # Packing list que corresponde a este romaneo (consignee + N° de venta)
    codigos = {_clean_code(c) for c in df[COL_CODIGO].dropna()} if COL_CODIGO in df.columns else set()
    total_cajas = int(round(df[COL_CAJAS].sum()))
    pl = match_packing_list(packing_lists or [], codigos, total_cajas, solo_uno=solo_uno)
    if not pl:
        raise ValueError('No se encontró el Packing List de este romaneo. '
                         'Subí su Packing List (de ahí salen el N° de venta y el consignee).')
    if not pl.get('venta'):
        raise ValueError(f"El Packing List {pl['archivo']} no tiene EXPORTER REFERENCE (N° de venta).")

    salida = _nombre_salida(pl['venta'])
    outpath = OUTPUT_FOLDER / salida
    shutil.copy(PLANILLA_BASE, outpath)

    wb = load_workbook(outpath)
    ws = wb[next(s for s in wb.sheetnames if 'PLANILLA' in s.upper())]

    for i, row_data in grouped.iterrows():
        excel_row = 34 + i
        if excel_row > 63:
            break
        ws[f'E{excel_row}'] = row_data['Descripcion']
        ws[f'F{excel_row}'] = round(float(row_data['Cajas']), 0)
        # Neto redondeado a kg enteros (0,5 redondea para arriba); el bruto queda con 2 decimales
        ws[f'G{excel_row}'] = int(math.floor(float(row_data['Peso_kg']) + 0.5))
        ws[f'L{excel_row}'] = round(float(row_data['Peso_Bruto']), 2)

    ws['E69'] = tropas_str
    ws['E70'] = lotes_str

    # CONSIGNEE (E25) desde el packing list
    # N° de Venta (secciones 5 y 6) desde EXPORTER REFERENCE del packing list
    venta = pl['venta'].strip()
    venta_val = int(venta) if venta.isdigit() else venta
    for celda in ('E32', 'E71'):
        ws[celda] = venta_val
        ws[celda].alignment = ws[celda].alignment.copy(horizontal='left')

    ws['E25'] = pl['consignee']
    ws['E25'].alignment = ws['E25'].alignment.copy(wrap_text=True, vertical='top')
    consignee_info = {
        'archivo': pl['archivo'],
        'nombre': pl['consignee'].splitlines()[0],
        'cajas_ok': pl['cajas'] == total_cajas,
        'cajas_pl': pl['cajas'],
    }

    wb.save(outpath)

    return {
        'nombre': filepath.name,
        'salida': salida,
        'cortes': len(grouped),
        'tropas': len(all_tropas),
        'lotes': len(all_lotes),
        'cajas': total_cajas,
        'consignee': consignee_info,
    }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    romaneo_files = request.files.getlist('romaneo_files')

    if not romaneo_files or all(f.filename == '' for f in romaneo_files):
        return jsonify({'error': 'No seleccionaste archivos ROMANEO'}), 400

    if not any(f.filename for f in request.files.getlist('packing_lists')):
        return jsonify({'error': 'Subí el Packing List de cada romaneo (de ahí salen el N° de venta y el consignee)'}), 400

    if not PLANILLA_BASE.exists():
        return jsonify({'error': 'No se encontró Planilla_BASE_.xlsx en el servidor'}), 500

    for f in UPLOAD_FOLDER.glob('*'):
        f.unlink()
    for f in OUTPUT_FOLDER.glob('*'):
        f.unlink()

    resultados = []
    errores = []

    # PACKING LISTs (PDF): de ahí sale el CONSIGNEE de cada carga
    packing_lists = []
    for f in request.files.getlist('packing_lists'):
        if not f.filename:
            continue
        pl_path = UPLOAD_FOLDER / f.filename
        f.save(pl_path)
        try:
            pl = parse_packing_list(pl_path)
            if not pl['consignee']:
                errores.append({'nombre': f.filename, 'error': 'No se encontró el bloque CONSIGNEE en el Packing List'})
                continue
            packing_lists.append(pl)
        except Exception as e:
            errores.append({'nombre': f.filename, 'error': f'Error leyendo Packing List: {e}'})

    romaneos_validos = [f for f in romaneo_files if f.filename]
    solo_uno = len(romaneos_validos) == 1

    for f in romaneos_validos:
        save_path = UPLOAD_FOLDER / f.filename
        f.save(save_path)
        try:
            r = process_file(save_path, packing_lists=packing_lists, solo_uno=solo_uno)
            resultados.append(r)
        except Exception as e:
            errores.append({'nombre': f.filename, 'error': str(e)})

    return jsonify({'resultados': resultados, 'errores': errores})


@app.route('/descargar/<filename>')
def descargar(filename):
    filepath = OUTPUT_FOLDER / filename
    if not filepath.exists():
        return 'Archivo no encontrado', 404
    return send_file(filepath, as_attachment=True)


@app.route('/descargar_todo')
def descargar_todo():
    archivos = list(OUTPUT_FOLDER.glob('PLANILLA OIA *.xlsx'))
    if not archivos:
        return 'No hay archivos para descargar', 404
    zip_path = OUTPUT_FOLDER / 'planillas_completadas.zip'
    with zipfile.ZipFile(zip_path, 'w') as zf:
        for f in archivos:
            zf.write(f, f.name)
    return send_file(zip_path, as_attachment=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
