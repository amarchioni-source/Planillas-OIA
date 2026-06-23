"""
app.py - Generador de Planillas Angus
Flask app para procesar archivos PIQUEO y generar planillas completadas
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
    Busca columnas por nombre ('Code' y 'Description'), ignorando mayúsculas.
    """
    xl = pd.ExcelFile(reporte_path)

    for sheet in xl.sheet_names:
        raw = pd.read_excel(reporte_path, sheet_name=sheet, header=None)

        # Buscar fila de encabezado que tenga 'code' y 'description'
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'code' in vals and 'description' in vals:
                df = pd.read_excel(reporte_path, sheet_name=sheet, header=i)
                df.columns = [str(c).strip() for c in df.columns]

                # Encontrar nombres exactos de columna (case-insensitive)
                col_code = next((c for c in df.columns if c.lower() == 'code'), None)
                col_desc = next((c for c in df.columns if c.lower() == 'description'), None)

                if col_code and col_desc:
                    mapping = {}
                    for _, r in df[[col_code, col_desc]].dropna(subset=[col_code]).iterrows():
                        code = str(r[col_code]).strip()
                        desc = str(r[col_desc]).strip()
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


def process_file(filepath, base_path, translation_dict=None):
    filepath = Path(filepath)
    outpath = OUTPUT_FOLDER / (filepath.stem + '_completada.xlsx')

    xl = pd.ExcelFile(filepath)
    data_sheet, header_row = find_data_sheet_and_header(filepath, xl)

    df = pd.read_excel(filepath, sheet_name=data_sheet, header=header_row)

    # Normalizar nombres de columna
    df.columns = [str(c).strip() for c in df.columns]

    # Filtro Angus: solo productos con AA en columna Producto
    df = df[df['Producto'].astype(str).str.contains(r'\bAA\b', na=False)]

    # Buscar columna de código por nombre (case-insensitive)
    col_codigo = next(
        (c for c in df.columns if c.lower() in ('código', 'codigo', 'code', 'cod', 'cod.')),
        None
    )

    grouped = (
        df.groupby('Producto')
        .agg(Cajas=('Cajas', 'sum'), Peso_kg=('Peso', 'sum'), Peso_Bruto=('Bruto', 'sum'))
        .reset_index()
    )

    # Si hay REPORTE DOC, agregar columna de código y traducir descripción
    if translation_dict and col_codigo:
        # Tomar el primer código asociado a cada Producto
        codigo_por_producto = (
            df.groupby('Producto')[col_codigo]
            .first()
            .reset_index()
        )
        codigo_por_producto.columns = ['Producto', 'Codigo']
        grouped = grouped.merge(codigo_por_producto, on='Producto', how='left')

        def get_description(row):
            codigo = str(row.get('Codigo', '')).strip()
            if codigo and codigo != 'nan' and codigo in translation_dict:
                return translation_dict[codigo]
            return row['Producto']  # fallback al original

        grouped['Descripcion'] = grouped.apply(get_description, axis=1)
    else:
        grouped['Descripcion'] = grouped['Producto']

    all_tropas = set()
    for t in df['Tropa'].dropna():
        all_tropas.update(extract_tropas(t))
    tropas_str = ', '.join(sorted(all_tropas, key=lambda x: int(x)))

    all_lotes = set()
    for d in df['Fecha P'].dropna():
        all_lotes.add(pd.Timestamp(d).strftime('%d/%m/%Y'))
    lotes_str = ', '.join(sorted(all_lotes))

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
        ws[f'E{excel_row}'] = row_data['Descripcion']   # descripción inglés (o original si no matchea)
        ws[f'F{excel_row}'] = round(float(row_data['Cajas']), 0)
        ws[f'G{excel_row}'] = round(float(row_data['Peso_kg']), 2)
        ws[f'L{excel_row}'] = round(float(row_data['Peso_Bruto']), 2)

    ws['E69'] = tropas_str
    ws['E70'] = lotes_str
    wb.save(outpath)

    # Contar cuántos matchearon
    matched = 0
    if translation_dict and 'Codigo' in grouped.columns:
        matched = grouped['Codigo'].apply(
            lambda c: str(c).strip() in translation_dict
        ).sum()

    return {
        'nombre': filepath.name,
        'salida': filepath.stem + '_completada.xlsx',
        'cortes': len(grouped),
        'tropas': len(all_tropas),
        'lotes': len(all_lotes),
        'base_note': base_note,
        'traducidos': int(matched),
        'sin_traduccion': int(len(grouped) - matched),
    }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    piqueo_files = request.files.getlist('piqueo_files')
    base_file = request.files.get('planilla_base')
    reporte_file = request.files.get('reporte_doc')

    if not piqueo_files or all(f.filename == '' for f in piqueo_files):
        return jsonify({'error': 'No seleccionaste archivos PIQUEO'}), 400

    if not reporte_file or reporte_file.filename == '':
        return jsonify({'error': 'Tenés que subir el REPORTE DOC para obtener las descripciones en inglés'}), 400

    # Limpiar carpetas
    for f in UPLOAD_FOLDER.glob('*'):
        f.unlink()
    for f in OUTPUT_FOLDER.glob('*'):
        f.unlink()

    # Guardar Planilla BASE
    if base_file and base_file.filename:
        base_path = UPLOAD_FOLDER / 'Planilla_BASE_.xlsx'
        base_file.save(base_path)
    elif PLANILLA_BASE.exists():
        base_path = PLANILLA_BASE
    else:
        return jsonify({'error': 'No se encontró Planilla_BASE_.xlsx'}), 400

    # Guardar y procesar REPORTE DOC
    reporte_path = UPLOAD_FOLDER / reporte_file.filename
    reporte_file.save(reporte_path)
    try:
        translation_dict = build_translation_dict(reporte_path)
    except Exception as e:
        return jsonify({'error': f'Error leyendo REPORTE DOC: {str(e)}'}), 400

    resultados = []
    errores = []

    for f in piqueo_files:
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
