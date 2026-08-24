from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# Archivo Excel donde se guardarán las devoluciones para las pruebas
EXCEL_FILE = 'devoluciones_registradas.xlsx'

def guardar_en_excel(datos_encabezado, lista_productos):
    """Guarda la devolución y sus productos en un archivo Excel persistente."""
    filas = []
    
    # Asociamos cada producto registrado con los datos generales del folio
    for prod in lista_productos:
        fila = {**datos_encabezado, **prod}
        filas.append(fila)
        
    df_nuevo = pd.DataFrame(filas)
    
    # Si el archivo ya existe, concatenamos; si no, lo creamos
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
        
    df_final.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/guardar', methods=['POST'])
def guardar():
    try:
        # 1. Extraemos los datos generales del formulario
        datos_encabezado = {
            'folio_devolucion': request.form.get('folio_devolucion'),
            'fecha': request.form.get('fecha'),
            'factura_pedido': request.form.get('factura_pedido'),
            'nombre_cliente': request.form.get('nombre_cliente'),
            'empresa': request.form.get('empresa'),
            'telefono': request.form.get('telefono'),
            'correo': request.form.get('correo'),
            'motivo': request.form.get('motivo'),
            'motivo_observaciones': request.form.get('motivo_observaciones'),
            'insp_completo': request.form.get('insp_completo'),
            'obs_completo': request.form.get('obs_completo'),
            'insp_empaque': request.form.get('insp_empaque'),
            'obs_empaque': request.form.get('obs_empaque'),
            'insp_accesorios': request.form.get('insp_accesorios'),
            'obs_accesorios': request.form.get('obs_accesorios'),
            'insp_dano': request.form.get('insp_dano'),
            'obs_dano': request.form.get('obs_dano'),
            'insp_garantia': request.form.get('insp_garantia'),
            'obs_garantia': request.form.get('obs_garantia'),
            'resolucion': request.form.get('resolucion'),
            'resolucion_observaciones': request.form.get('resolucion_observaciones'),
            'costo_devolucion': request.form.get('costo_devolucion'),
            'area_responsable': request.form.get('area_responsable'),
            'fecha_cierre': request.form.get('fecha_cierre'),
            'accion_correctiva': request.form.get('accion_correctiva'),
            'fecha_registro_sistema': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        # 2. Extraemos las listas de la tabla dinámica de productos
        codigos = request.form.getlist('prod_codigo[]')
        descripciones = request.form.getlist('prod_descripcion[]')
        clasificaciones = request.form.getlist('prod_clasificacion[]')
        cants_compradas = request.form.getlist('prod_cant_comprada[]')
        cants_devueltas = request.form.getlist('prod_cant_devuelta[]')
        precios = request.form.getlist('prod_precio[]')
        totales = request.form.getlist('prod_total[]')

        lista_productos = []
        for i in range(len(codigos)):
            lista_productos.append({
                'prod_codigo': codigos[i],
                'prod_descripcion': descripciones[i],
                'prod_clasificacion': clasificaciones[i],
                'prod_cant_comprada': cants_compradas[i],
                'prod_cant_devuelta': cants_devueltas[i],
                'prod_precio_unitario': precios[i],
                'prod_total': totales[i]
            })

        # 3. Guardamos en el archivo Excel
        guardar_en_excel(datos_encabezado, lista_productos)

        return jsonify({'status': 'ok', 'mensaje': '¡Devolución registrada exitosamente!'})

    except Exception as e:
        return jsonify({'status': 'error', 'mensaje': str(e)}), 500

@app.route('/descargar-excel')
def descargar_excel():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    return "Aún no hay registros guardados.", 404

if __name__ == '__main__':
    app.run(debug=True)