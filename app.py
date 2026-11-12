"""
app.py - Generador de Planillas Angus
Flask app para procesar archivos ROMANEO y generar planillas completadas
"""
import os, re, shutil, zipfile
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


def build_translation_dict(reporte_path):
    """
    Lee el REPORTE DOC y construye un diccionario {código: descripción_inglés}.
    Lee solo 20 filas para encontrar el encabezado, luego carga únicamente
    las dos columnas necesarias para minimizar uso de RAM.
    """
    xl = pd.ExcelFile(reporte_path)

    for sheet in xl.sheet_names:
        raw = pd.read_excel(reporte_path, sheet_name=sheet, header=None, nrows=20)

        header_row = None
        col_code_idx = None
        col_desc_idx = None
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'code' in vals and 'description' in vals:
                header_row = i
                col_code_idx = vals.index('code')
                col_desc_idx = vals.index('description')
                break

        if header_row is None:
            continue

        df = pd.read_excel(
            reporte_path,
            sheet_name=sheet,
            header=header_row,
            usecols=[col_code_idx, col_desc_idx]
        )
        df.columns = [str(c).strip() for c in df.columns]

        col_code = next((c for c in df.columns if c.lower() == 'code'), None)
        col_desc = next((c for c in df.columns if c.lower() == 'description'), None)

        if col_code and col_desc:
            mapping = {}
            df = df.dropna(subset=[col_code]).reset_index(drop=True)
            for idx in range(len(df)):
                code = str(df[col_code].iloc[idx]).strip()
                desc = str(df[col_desc].iloc[idx]).strip()
                if code and desc and code != 'nan':
                    mapping[code] = desc
            return mapping

    raise ValueError(
        "No se encontraron columnas 'Code' y 'Description' en el REPORTE DOC."
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
COL_OLD_CODE = 'Old Code'
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


def process_file(filepath, base_path, translation_dict=None):
    filepath = Path(filepath)
    outpath = OUTPUT_FOLDER / (filepath.stem + '_completada.xlsx')

    xl = pd.ExcelFile(filepath)
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

    # Descripción en inglés: 1) REPORTE DOC por Cod Prod / Old Code,
    # 2) columna 'Product' del romaneo, 3) Producto original.
    first = df.groupby(COL_PRODUCTO).first()

    def get_description(producto):
        row = first.loc[producto]
        if translation_dict:
            for col in (COL_CODIGO, COL_OLD_CODE):
                if col in row.index:
                    code = _clean_code(row[col])
                    if code and code in translation_dict:
                        return translation_dict[code], True
        if COL_PRODUCT_EN in row.index:
            en = str(row[COL_PRODUCT_EN]).strip()
            if en and en.lower() != 'nan':
                return en, False
        return producto, False

    descs = grouped['Producto'].apply(get_description)
    grouped['Descripcion'] = descs.apply(lambda t: t[0])
    grouped['EnReporte'] = descs.apply(lambda t: t[1])

    all_tropas = set()
    for t in df[COL_TROPA].dropna():
        all_tropas.update(extract_tropas(t))
    tropas_str = ', '.join(sorted(all_tropas, key=lambda x: int(x)))

    all_lotes = set()
    for d in df[COL_FECHA_PROD].dropna():
        all_lotes.add(pd.Timestamp(d).normalize())
    lotes_str = ', '.join(d.strftime('%d/%m/%Y') for d in sorted(all_lotes))

    has_planilla = any('PLANILLA' in s.upper() for s in xl.sheet_names)
    if has_planilla:
        shutil.copy(filepath, outpath)
        base_note = ''
    else:
        shutil.copy(base_path, outpath)
        base_note = ' [usó Planilla_BASE_]'

    wb = load_workbook(outpath)
    ws = wb[next(s for s in wb.sheetnames if 'PLANILLA' in s.upper())]

    for i, row_data in grouped.iterrows():
        excel_row = 34 + i
        if excel_row > 63:
            break
        ws[f'E{excel_row}'] = row_data['Descripcion']
        ws[f'F{excel_row}'] = round(float(row_data['Cajas']), 0)
        ws[f'G{excel_row}'] = round(float(row_data['Peso_kg']), 2)
        ws[f'L{excel_row}'] = round(float(row_data['Peso_Bruto']), 2)

    ws['E69'] = tropas_str
    ws['E70'] = lotes_str
    wb.save(outpath)

    matched = int(grouped['EnReporte'].sum()) if translation_dict else 0

    return {
        'nombre': filepath.name,
        'salida': filepath.stem + '_completada.xlsx',
        'cortes': len(grouped),
        'tropas': len(all_tropas),
        'lotes': len(all_lotes),
        'base_note': base_note,
        'traducidos': int(matched),
        'sin_traduccion': int(len(grouped) - matched) if translation_dict else 0,
    }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    romaneo_files = request.files.getlist('romaneo_files')
    base_file = request.files.get('planilla_base')
    reporte_file = request.files.get('reporte_doc')

    if not romaneo_files or all(f.filename == '' for f in romaneo_files):
        return jsonify({'error': 'No seleccionaste archivos ROMANEO'}), 400

    for f in UPLOAD_FOLDER.glob('*'):
        f.unlink()
    for f in OUTPUT_FOLDER.glob('*'):
        f.unlink()

    if base_file and base_file.filename:
        base_path = UPLOAD_FOLDER / 'Planilla_BASE_.xlsx'
        base_file.save(base_path)
    elif PLANILLA_BASE.exists():
        base_path = PLANILLA_BASE
    else:
        return jsonify({'error': 'No se encontró Planilla_BASE_.xlsx'}), 400

    # REPORTE DOC opcional: si no viene, se usa la columna 'Product' del romaneo
    translation_dict = None
    if reporte_file and reporte_file.filename:
        reporte_path = UPLOAD_FOLDER / reporte_file.filename
        reporte_file.save(reporte_path)
        try:
            translation_dict = build_translation_dict(reporte_path)
        except Exception as e:
            return jsonify({'error': f'Error leyendo REPORTE DOC: {str(e)}'}), 400

    resultados = []
    errores = []

    for f in romaneo_files:
        if f.filename == '':
            continue
        save_path = UPLOAD_FOLDER / f.filename
        f.save(save_path)
        try:
            r = process_file(save_path, base_path, translation_dict)
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
    archivos = list(OUTPUT_FOLDER.glob('*_completada.xlsx'))
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
