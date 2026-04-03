from flask import Flask, render_template, request, redirect, session, jsonify, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import traceback
import io
import os
import json

app = Flask(__name__)

# Configuración de secret key - Usar variable de entorno en producción
app.secret_key = os.environ.get('SECRET_KEY', 'clave_secreta_para_desarrollo_123456')

# 🔗 CONFIGURACIÓN GOOGLE SHEETS
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

# Detectar si estamos en producción (Render)
IN_PRODUCTION = os.environ.get('RENDER', False)

try:
    if IN_PRODUCTION:
        # En Render, las credenciales vienen como variable de entorno
        creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if creds_json:
            creds_dict = json.loads(creds_json)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            raise Exception("No se encontraron credenciales de Google")
    else:
        # En desarrollo local
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
    
    client = gspread.authorize(creds)
    spreadsheet = client.open("Respuestas Formulario")
    
    prof_sheet = spreadsheet.worksheet("PROFESORES")
    est_sheet = spreadsheet.worksheet("ESTUDIANTES")
    resp_sheet = spreadsheet.worksheet("RESPUESTAS")
    mat_sheet = spreadsheet.worksheet("MATERIAS")
    config_sheet = spreadsheet.worksheet("CONFIG")
    
    print("✅ Conexión exitosa con Google Sheets")
    
except Exception as e:
    print(f"❌ Error de conexión: {e}")
    if not IN_PRODUCTION:
        raise e

# ==================== FUNCIONES AUXILIARES ====================

def convertir_a_booleano(valor):
    """Convierte cualquier valor a booleano"""
    if valor is None:
        return False
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, str):
        valor_limpio = valor.strip().upper()
        return valor_limpio == "TRUE" or valor_limpio == "VERDADERO" or valor_limpio == "1"
    if isinstance(valor, (int, float)):
        return valor == 1
    return False

def booleano_a_texto(valor):
    """Convierte booleano a texto"""
    return "TRUE" if valor else "FALSE"

def verificar_fecha_valida():
    """Verifica si la fecha actual está dentro del período permitido"""
    try:
        config_data = config_sheet.get_all_records()
        for row in config_data:
            if row.get('clave') == 'activo':
                activo = convertir_a_booleano(row.get('valor', False))
            elif row.get('clave') == 'fecha_inicio':
                fecha_inicio_str = row.get('valor', '')
            elif row.get('clave') == 'fecha_fin':
                fecha_fin_str = row.get('valor', '')
        
        # Si no está activo, rechazar acceso
        if not activo:
            return False, "El sistema está desactivado"
        
        # Verificar fechas
        ahora = datetime.now()
        
        if fecha_inicio_str:
            try:
                fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d %H:%M:%S")
                if ahora < fecha_inicio:
                    return False, f"El sistema estará disponible a partir del {fecha_inicio.strftime('%d/%m/%Y %H:%M')}"
            except:
                pass
        
        if fecha_fin_str:
            try:
                fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d %H:%M:%S")
                if ahora > fecha_fin:
                    return False, f"El sistema expiró el {fecha_fin.strftime('%d/%m/%Y %H:%M')}"
            except:
                pass
        
        return True, "Sistema activo"
    
    except Exception as e:
        print(f"Error verificando fecha: {e}")
        return True, "No se pudo verificar la fecha"  # Permitir acceso por defecto si hay error

def incrementar_contador_descargas(profesor):
    """Incrementa el contador de descargas para un profesor"""
    try:
        # Buscar o crear hoja de estadísticas
        try:
            stats_sheet = spreadsheet.worksheet("ESTADISTICAS")
        except:
            # Crear la hoja si no existe
            stats_sheet = spreadsheet.add_worksheet("ESTADISTICAS", rows=100, cols=10)
            stats_sheet.append_row(["profesor", "descargas_pdf", "ultima_descarga"])
        
        # Buscar al profesor
        registros = stats_sheet.get_all_records()
        fila_encontrada = None
        for idx, registro in enumerate(registros, start=2):
            if registro.get('profesor', '').upper() == profesor.upper():
                fila_encontrada = idx
                break
        
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if fila_encontrada:
            # Actualizar contador existente
            descargas_actual = int(stats_sheet.cell(fila_encontrada, 2).value or 0)
            stats_sheet.update_cell(fila_encontrada, 2, descargas_actual + 1)
            stats_sheet.update_cell(fila_encontrada, 3, ahora)
        else:
            # Crear nuevo registro
            stats_sheet.append_row([profesor.upper(), 1, ahora])
        
        return True
    except Exception as e:
        print(f"Error actualizando contador: {e}")
        return False

def generar_reporte_pdf(profesor, cursos_data, materias_data, solo_marcadas=True):
    """Genera reporte PDF solo de materias marcadas"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter),
                           rightMargin=30, leftMargin=30,
                           topMargin=30, bottomMargin=30)
    
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        'TituloStyle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=1,
        spaceAfter=20,
        textColor=colors.HexColor('#4a6cf7')
    )
    
    subtitulo_style = ParagraphStyle(
        'SubtituloStyle',
        parent=styles['Heading2'],
        fontSize=12,
        alignment=1,
        spaceAfter=10,
        textColor=colors.HexColor('#666666')
    )
    
    curso_style = ParagraphStyle(
        'CursoStyle',
        parent=styles['Heading3'],
        fontSize=14,
        spaceAfter=10,
        spaceBefore=15,
        textColor=colors.HexColor('#333333')
    )
    
    elementos = []
    
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    titulo = Paragraph(f"<b>REPORTE DE EVALUACIONES</b>", titulo_style)
    elementos.append(titulo)
    
    subtitulo = Paragraph(f"Profesor: {profesor}<br/>Fecha: {fecha_actual}<br/>Mostrando: {'Solo materias evaluadas' if solo_marcadas else 'Todas las materias'}", subtitulo_style)
    elementos.append(subtitulo)
    elementos.append(Spacer(1, 20))
    
    total_general_marcadas = 0
    total_general_posibles = 0
    
    for curso_nombre, alumnos in cursos_data.items():
        elementos.append(Paragraph(f"<b>Curso: {curso_nombre}</b>", curso_style))
        elementos.append(Spacer(1, 10))
        
        # Filtrar materias que tienen al menos una evaluación marcada si solo_marcadas es True
        materias_a_mostrar = {}
        if solo_marcadas:
            # Identificar qué materias tienen al menos un checkbox marcado
            for materia_id in sorted(materias_data.keys()):
                tiene_marcada = False
                for alumno in alumnos:
                    key = f"{curso_nombre}_{alumno}_{materia_id}"
                    if session.get(f"estado_temp_{key}", False):
                        tiene_marcada = True
                        break
                if tiene_marcada:
                    materias_a_mostrar[materia_id] = materias_data[materia_id]
        else:
            materias_a_mostrar = materias_data
        
        if not materias_a_mostrar:
            elementos.append(Paragraph("<i>No hay materias evaluadas en este curso</i>", styles['Italic']))
            elementos.append(Spacer(1, 20))
            continue
        
        encabezados = ["Alumno"] + list(materias_a_mostrar.values())
        tabla_datos = [encabezados]
        
        total_materias_marcadas = 0
        total_posibles = 0
        
        for alumno in alumnos:
            fila = [alumno]
            tiene_alguna_marcada = False
            
            for materia_id in sorted(materias_a_mostrar.keys()):
                key = f"{curso_nombre}_{alumno}_{materia_id}"
                estado = session.get(f"estado_temp_{key}", False)
                
                if estado:
                    fila.append("✓")
                    total_materias_marcadas += 1
                    tiene_alguna_marcada = True
                else:
                    fila.append("")
                total_posibles += 1
            
            # Si es modo "solo marcadas" y el alumno no tiene ninguna marcada, omitirlo
            if solo_marcadas and not tiene_alguna_marcada:
                continue
                
            tabla_datos.append(fila)
        
        if len(tabla_datos) <= 1:  # Solo encabezados, sin alumnos
            elementos.append(Paragraph("<i>No hay alumnos con evaluaciones en este curso</i>", styles['Italic']))
            elementos.append(Spacer(1, 20))
            continue
        
        tabla = Table(tabla_datos, repeatRows=1)
        
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4a6cf7')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elementos.append(tabla)
        
        if total_posibles > 0:
            porcentaje = (total_materias_marcadas / total_posibles) * 100
            estadisticas = Paragraph(
                f"<i>Estadísticas: {total_materias_marcadas} evaluaciones de {total_posibles} posibles ({porcentaje:.1f}%)</i>",
                styles['Italic']
            )
            elementos.append(Spacer(1, 5))
            elementos.append(estadisticas)
            total_general_marcadas += total_materias_marcadas
            total_general_posibles += total_posibles
        
        elementos.append(Spacer(1, 20))
    
    elementos.append(Spacer(1, 30))
    
    if total_general_posibles > 0:
        porcentaje_general = (total_general_marcadas / total_general_posibles) * 100
        
        resumen_style = ParagraphStyle(
            'ResumenStyle',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            textColor=colors.HexColor('#666666')
        )
        
        resumen_texto = f"""
        <b>RESUMEN GENERAL</b><br/>
        Total evaluaciones realizadas: {total_general_marcadas}<br/>
        Total evaluaciones posibles: {total_general_posibles}<br/>
        Porcentaje de avance: {porcentaje_general:.1f}%<br/>
        Generado: {fecha_actual}
        """
        
        resumen = Paragraph(resumen_texto, resumen_style)
        elementos.append(resumen)
    
    doc.build(elementos)
    buffer.seek(0)
    
    return buffer

# ==================== RUTAS ====================

@app.route('/', methods=['GET', 'POST'])
def login():
    # Verificar fecha de caducidad antes de mostrar login
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip().upper()
        
        if not usuario:
            return "Por favor ingrese un usuario"
        
        try:
            profesores = prof_sheet.get_all_records()
            
            for p in profesores:
                if p.get('usuario', '').upper() == usuario:
                    materias = []
                    for campo in ['m1', 'm2', 'm3']:
                        if p.get(campo):
                            try:
                                materias.append(int(p[campo]))
                            except:
                                pass
                    
                    cursos_str = str(p.get('cursos', ''))
                    cursos = [x.strip() for x in cursos_str.split(",") if x.strip()]
                    
                    # Obtener nombre completo del profesor
                    nombre_completo = p.get('nombre_completo', usuario)
                    
                    session['usuario'] = usuario
                    session['nombre_completo'] = nombre_completo
                    session['materias'] = materias
                    session['cursos'] = cursos
                    
                    return redirect('/panel')
            
            return f"❌ Usuario '{usuario}' no encontrado"
        
        except Exception as e:
            return f"Error al verificar usuario: {e}"
    
    return render_template('login.html')

@app.route('/panel')
def panel():
    if 'usuario' not in session:
        return redirect('/')
    
    # Verificar fecha de caducidad
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    try:
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        
        for est in estudiantes:
            curso = str(est.get('curso', ''))
            nombre = est.get('nombre', '')
            if curso in session['cursos'] and nombre:
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso]:
                    cursos[curso].append(nombre)
        
        materias_data = mat_sheet.get_all_records()
        materias = {}
        for m in materias_data:
            try:
                materias[int(m['id'])] = m['nombre']
            except:
                pass
        
        todas_respuestas = resp_sheet.get_all_records()
        estado = {}
        
        for respuesta in todas_respuestas:
            if respuesta.get('profesor', '').upper() == session['usuario'].upper():
                curso = respuesta.get('curso', '')
                alumno = respuesta.get('alumno', '')
                
                for i in range(1, 16):
                    key = f"{curso}_{alumno}_{i}"
                    columna = f"m{i}"
                    valor = respuesta.get(columna, False)
                    estado[key] = convertir_a_booleano(valor)
        
        for key, value in estado.items():
            session[f"estado_temp_{key}"] = value
        
        return render_template('panel.html',
                               cursos=cursos,
                               materias=materias,
                               materias_user=session['materias'],
                               estado=estado,
                               usuario=session['usuario'],
                               nombre_completo=session.get('nombre_completo', session['usuario']))
    
    except Exception as e:
        return f"Error al cargar el panel: {e}"

@app.route('/guardar', methods=['POST'])
def guardar():
    # Verificar fecha de caducidad
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return jsonify({
            "success": False,
            "error": mensaje
        }), 403
    
    try:
        data = request.json
        
        profesor = session.get('usuario', '').upper()
        curso = data.get('curso')
        alumno = data.get('alumno')
        materia = int(data.get('materia'))
        valor = data.get('valor')
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        valor_texto = booleano_a_texto(valor)
        
        key = f"{curso}_{alumno}_{materia}"
        session[f"estado_temp_{key}"] = valor
        
        todas_filas = resp_sheet.get_all_values()
        
        num_fila = None
        for idx, fila in enumerate(todas_filas, start=1):
            if idx == 1:
                continue
            if len(fila) >= 3:
                if (fila[0].upper() == profesor and 
                    fila[1] == curso and 
                    fila[2] == alumno):
                    num_fila = idx
                    break
        
        columna_materia = 3 + materia
        
        if num_fila:
            resp_sheet.update_cell(num_fila, columna_materia, valor_texto)
            resp_sheet.update_cell(num_fila, 19, fecha)
        else:
            nueva_fila = [profesor, curso, alumno]
            for _ in range(15):
                nueva_fila.append("FALSE")
            nueva_fila.append(fecha)
            
            resp_sheet.append_row(nueva_fila)
            
            todas_filas_nuevas = resp_sheet.get_all_values()
            for idx, fila in enumerate(todas_filas_nuevas, start=1):
                if idx == 1:
                    continue
                if len(fila) >= 3:
                    if (fila[0].upper() == profesor and 
                        fila[1] == curso and 
                        fila[2] == alumno):
                        resp_sheet.update_cell(idx, columna_materia, valor_texto)
                        break
        
        return jsonify({
            "success": True,
            "mensaje": "Guardado correctamente"
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/pdf')
def pdf():
    if 'usuario' not in session:
        return redirect('/')
    
    # Verificar fecha de caducidad
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    try:
        profesor = session['usuario']
        
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        
        for est in estudiantes:
            curso = str(est.get('curso', ''))
            nombre = est.get('nombre', '')
            if curso in session['cursos'] and nombre:
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso]:
                    cursos[curso].append(nombre)
        
        materias_data = mat_sheet.get_all_records()
        materias = {}
        for m in materias_data:
            try:
                materias[int(m['id'])] = m['nombre']
            except:
                pass
        
        materias_profesor = {mid: materias[mid] for mid in session['materias'] if mid in materias}
        
        # Generar PDF solo con materias marcadas
        pdf_buffer = generar_reporte_pdf(profesor, cursos, materias_profesor, solo_marcadas=True)
        
        # Incrementar contador de descargas
        incrementar_contador_descargas(profesor)
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=f"reporte_{profesor}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
    
    except Exception as e:
        return f"Error al generar PDF: {e}"

@app.route('/logout')
def logout():
    keys_to_remove = [key for key in session.keys() if key.startswith('estado_temp_')]
    for key in keys_to_remove:
        session.pop(key, None)
    
    session.clear()
    return redirect('/')

@app.route('/diagnostico')
def diagnostico():
    if 'usuario' not in session:
        return redirect('/')
    
    try:
        todas_filas = resp_sheet.get_all_values()
        
        resultado = {
            "usuario": session['usuario'],
            "encabezados": todas_filas[0] if todas_filas else [],
            "mis_registros": []
        }
        
        for i, fila in enumerate(todas_filas[1:], start=2):
            if len(fila) > 0 and fila[0].upper() == session['usuario'].upper():
                registro = {
                    "fila_numero": i,
                    "profesor": fila[0] if len(fila) > 0 else "",
                    "curso": fila[1] if len(fila) > 1 else "",
                    "alumno": fila[2] if len(fila) > 2 else "",
                }
                for m in range(1, 16):
                    col_idx = 3 + m - 1
                    registro[f"m{m}"] = fila[col_idx] if len(fila) > col_idx else ""
                registro["fecha"] = fila[18] if len(fila) > 18 else ""
                resultado["mis_registros"].append(registro)
        
        return jsonify(resultado)
    
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/test_conexion')
def test_conexion():
    return jsonify({"success": True, "mensaje": "Servidor activo"})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)