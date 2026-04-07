from flask import Flask, render_template, request, redirect, session, jsonify, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import traceback
import io
import os
import json
import unicodedata

app = Flask(__name__)

# Configuración de secret key
app.secret_key = os.environ.get('SECRET_KEY', 'clave_secreta_para_desarrollo_123456')

# Configuración Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

IN_PRODUCTION = os.environ.get('RENDER', False)

# ==================== FUNCIÓN DE NORMALIZACIÓN SOLO PARA COMPARACIÓN ====================
def normalizar_texto(texto):
    """Normaliza texto SOLO para comparación (no modifica lo que se guarda o muestra)"""
    if not texto:
        return ""
    texto = str(texto).strip().upper()
    texto = ' '.join(texto.split())
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) 
                    if unicodedata.category(c) != 'Mn')
    return texto
# =======================================================================================

try:
    if IN_PRODUCTION:
        creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if creds_json:
            creds_dict = json.loads(creds_json)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            raise Exception("No se encontraron credenciales de Google")
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
    
    client = gspread.authorize(creds)
    spreadsheet = client.open("Respuestas Formulario")
    
    prof_sheet = spreadsheet.worksheet("PROFESORES")
    est_sheet = spreadsheet.worksheet("ESTUDIANTES")
    resp_sheet = spreadsheet.worksheet("RESPUESTAS")
    mat_sheet = spreadsheet.worksheet("MATERIAS")
    
    try:
        config_sheet = spreadsheet.worksheet("CONFIG")
        config_sheet.update('B2', 'TRUE')
    except:
        config_sheet = spreadsheet.add_worksheet("CONFIG", rows=10, cols=2)
        config_sheet.append_row(["clave", "valor"])
        config_sheet.append_row(["activo", "TRUE"])
        config_sheet.append_row(["fecha_inicio", "2026-01-01 00:00:00"])
        config_sheet.append_row(["fecha_fin", "2026-12-31 23:59:59"])
    
    try:
        stats_sheet = spreadsheet.worksheet("ESTADISTICAS")
        headers = stats_sheet.row_values(1)
        if len(headers) < 3 or headers[0] != 'profesor':
            stats_sheet.update('A1:C1', [['profesor', 'descargas_pdf', 'ultima_descarga']])
    except:
        stats_sheet = spreadsheet.add_worksheet("ESTADISTICAS", rows=100, cols=10)
        stats_sheet.append_row(["profesor", "descargas_pdf", "ultima_descarga"])
    
    print("✅ Conexión exitosa con Google Sheets")
    
except Exception as e:
    print(f"❌ Error de conexión: {e}")
    if not IN_PRODUCTION:
        raise e

# ==================== FUNCIONES AUXILIARES ====================

def convertir_a_booleano(valor):
    if valor is None:
        return False
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return valor == 1 or valor == True
    if isinstance(valor, str):
        valor_limpio = valor.strip().upper()
        if valor_limpio in ['TRUE', 'VERDADERO', '1', 'YES', 'SI', 'X', '✓']:
            return True
        return False
    return False

def booleano_a_texto(valor):
    return "TRUE" if valor else "FALSE"

def verificar_fecha_valida():
    try:
        config_data = config_sheet.get_all_records()
        activo = True
        fecha_inicio_str = ""
        fecha_fin_str = ""
        
        for row in config_data:
            clave = row.get('clave', '')
            valor = row.get('valor', '')
            if clave == 'activo':
                activo = convertir_a_booleano(valor)
            elif clave == 'fecha_inicio':
                fecha_inicio_str = valor
            elif clave == 'fecha_fin':
                fecha_fin_str = valor
        
        if not activo:
            return False, "El sistema está desactivado"
        
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
        return True, "No se pudo verificar la fecha"

def incrementar_contador_descargas(profesor):
    try:
        stats_sheet = spreadsheet.worksheet("ESTADISTICAS")
        registros = stats_sheet.get_all_records()
        fila_encontrada = None
        for idx, registro in enumerate(registros, start=2):
            if registro.get('profesor', '').upper() == profesor.upper():
                fila_encontrada = idx
                break
        
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if fila_encontrada:
            descargas_actual = int(stats_sheet.cell(fila_encontrada, 2).value or 0)
            stats_sheet.update_cell(fila_encontrada, 2, descargas_actual + 1)
            stats_sheet.update_cell(fila_encontrada, 3, ahora)
            return True
        else:
            stats_sheet.append_row([profesor.upper(), 1, ahora])
            return True
        
    except Exception as e:
        print(f"❌ Error actualizando contador: {e}")
        return False

def obtener_materias_por_curso(profesor_dict, cursos_disponibles):
    materias_por_curso = {curso: [] for curso in cursos_disponibles}
    
    for i in range(1, 4):
        materia_id = profesor_dict.get(f'm{i}')
        cursos_str = profesor_dict.get(f'cursos_m{i}', '')
        
        if not materia_id or str(materia_id).strip() == '':
            continue
        
        try:
            materia_id_int = int(float(materia_id))
            if materia_id_int <= 0 or materia_id_int > 15:
                continue
        except:
            continue
        
        if cursos_str and str(cursos_str).strip():
            cursos_aplicacion = [c.strip() for c in str(cursos_str).split(',') if c.strip()]
            
            for curso in cursos_aplicacion:
                if curso in materias_por_curso:
                    if materia_id_int not in materias_por_curso[curso]:
                        materias_por_curso[curso].append(materia_id_int)
    
    for curso in materias_por_curso:
        materias_por_curso[curso].sort()
    
    return materias_por_curso

def generar_reporte_pdf(profesor, cursos_data, todas_materias, materias_por_curso, solo_marcadas=True, nombre_completo=""):
    """Genera reporte PDF con búsqueda inteligente en sesión (primero original, luego normalizado)"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
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
    
    nombre_mostrar = nombre_completo if nombre_completo else profesor
    subtitulo = Paragraph(f"Profesor: {nombre_mostrar}<br/>Usuario: {profesor}<br/>Fecha: {fecha_actual}<br/>Mostrando: {'Solo materias evaluadas' if solo_marcadas else 'Todas las materias'}", subtitulo_style)
    elementos.append(subtitulo)
    elementos.append(Spacer(1, 20))
    
    total_general_marcadas = 0
    total_general_posibles = 0
    
    for curso_nombre, alumnos in cursos_data.items():
        elementos.append(Paragraph(f"<b>Curso: {curso_nombre}</b>", curso_style))
        elementos.append(Spacer(1, 10))
        
        materias_curso_ids = materias_por_curso.get(curso_nombre, [])
        
        if not materias_curso_ids:
            elementos.append(Paragraph("<i>No hay materias asignadas para este curso</i>", styles['Italic']))
            elementos.append(Spacer(1, 20))
            continue
        
        materias_a_mostrar = {}
        for materia_id in materias_curso_ids:
            if materia_id in todas_materias and todas_materias[materia_id]:
                if solo_marcadas:
                    tiene_marcada = False
                    for alumno in alumnos:
                        # PRIMERO buscar con clave ORIGINAL (como se guarda en sesión)
                        key_original = f"{curso_nombre}_{alumno}_{materia_id}"
                        estado_valor = session.get(f"estado_temp_{key_original}", False)
                        
                        # Si no encuentra, buscar con NORMALIZADO (compatibilidad hacia atrás)
                        if not estado_valor:
                            alumno_norm = normalizar_texto(alumno)
                            curso_norm = normalizar_texto(curso_nombre)
                            key_norm = f"{curso_norm}_{alumno_norm}_{materia_id}"
                            estado_valor = session.get(f"estado_temp_{key_norm}", False)
                        
                        if estado_valor:
                            tiene_marcada = True
                            break
                    if tiene_marcada:
                        materias_a_mostrar[materia_id] = todas_materias[materia_id]
                else:
                    materias_a_mostrar[materia_id] = todas_materias[materia_id]
        
        if not materias_a_mostrar:
            elementos.append(Paragraph("<i>No hay evaluaciones marcadas en este curso</i>", styles['Italic']))
            elementos.append(Spacer(1, 20))
            continue
        
        encabezados = ["Alumno"] + list(materias_a_mostrar.values())
        tabla_datos = [encabezados]
        
        total_materias_marcadas = 0
        total_posibles = 0
        
        for alumno in alumnos:
            fila = [alumno]
            tiene_alguna_marcada = False
            
            for materia_id in materias_a_mostrar.keys():
                # PRIMERO buscar con clave ORIGINAL
                key_original = f"{curso_nombre}_{alumno}_{materia_id}"
                estado_valor = session.get(f"estado_temp_{key_original}", False)
                
                # Si no encuentra, buscar con NORMALIZADO
                if not estado_valor:
                    alumno_norm = normalizar_texto(alumno)
                    curso_norm = normalizar_texto(curso_nombre)
                    key_norm = f"{curso_norm}_{alumno_norm}_{materia_id}"
                    estado_valor = session.get(f"estado_temp_{key_norm}", False)
                
                if estado_valor:
                    fila.append("✓")
                    total_materias_marcadas += 1
                    tiene_alguna_marcada = True
                else:
                    fila.append("")
                total_posibles += 1
            
            if solo_marcadas and not tiene_alguna_marcada:
                continue
            
            tabla_datos.append(fila)
        
        if len(tabla_datos) <= 1:
            elementos.append(Paragraph("<i>No hay alumnos con evaluaciones en este curso</i>", styles['Italic']))
            elementos.append(Spacer(1, 20))
            continue
        
        tabla = Table(tabla_datos, repeatRows=1)
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4a6cf7')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
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
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip().upper()
        
        if not usuario:
            return "Por favor ingrese un usuario"
        
        try:
            todas_filas = prof_sheet.get_all_values()
            
            if not todas_filas or len(todas_filas) < 2:
                return "Error: No hay datos en la hoja PROFESORES"
            
            encabezados = todas_filas[0]
            
            profesor_encontrado = None
            fila_datos = None
            
            for fila in todas_filas[1:]:
                if len(fila) > 0 and fila[0] and fila[0].strip().upper() == usuario:
                    profesor_encontrado = True
                    fila_datos = fila
                    break
            
            if profesor_encontrado and fila_datos:
                profesor_dict = {}
                for idx, header in enumerate(encabezados):
                    if header and str(header).strip():
                        if idx < len(fila_datos):
                            profesor_dict[str(header).strip()] = fila_datos[idx]
                        else:
                            profesor_dict[str(header).strip()] = ''
                
                cursos_set = set()
                for i in range(1, 4):
                    cursos_str = profesor_dict.get(f'cursos_m{i}', '')
                    if cursos_str:
                        for c in str(cursos_str).split(','):
                            curso_limpio = c.strip()
                            if curso_limpio:
                                cursos_set.add(curso_limpio)
                
                cursos = sorted(list(cursos_set))
                materias_por_curso = obtener_materias_por_curso(profesor_dict, cursos)
                
                nombre_completo = profesor_dict.get('nombre_completo', usuario)
                if not nombre_completo or str(nombre_completo).strip() == '':
                    nombre_completo = usuario
                
                session['usuario'] = usuario
                session['nombre_completo'] = str(nombre_completo)
                session['cursos'] = cursos
                session['materias_por_curso'] = materias_por_curso
                
                print(f"✅ Login exitoso: {usuario} - {nombre_completo}")
                print(f"   Cursos detectados: {cursos}")
                
                return redirect('/panel')
            
            return render_template('error.html', mensaje=f"Usuario '{usuario}' no encontrado")
        
        except Exception as e:
            print(f"Error en login: {e}")
            print(traceback.format_exc())
            return f"Error al verificar usuario: {e}"
    
    return render_template('login.html')

@app.route('/panel')
def panel():
    if 'usuario' not in session:
        return redirect('/')
    
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    try:
        usuario = session['usuario']
        cursos_profesor = session.get('cursos', [])
        materias_por_curso = session.get('materias_por_curso', {})
        
        print(f"🔍 Panel para: {usuario}")
        
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            
            if curso and nombre and curso in cursos_profesor:
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso] and nombre:
                    cursos[curso].append(nombre)
        
        materias_data = mat_sheet.get_all_records()
        todas_materias = {}
        for m in materias_data:
            try:
                id_materia = int(float(m.get('id', 0)))
                nombre_materia = str(m.get('nombre', '')).strip()
                if id_materia > 0 and nombre_materia and nombre_materia != '':
                    todas_materias[id_materia] = nombre_materia
            except Exception as e:
                continue
        
        todas_respuestas = resp_sheet.get_all_records()
        estado = {}
        contador_cargados = 0
        
        print(f"📥 Cargando respuestas previas para {usuario}...")
        
        for respuesta in todas_respuestas:
            profesor_resp = respuesta.get('profesor', '')
            if profesor_resp and str(profesor_resp).strip().upper() == usuario:
                curso = respuesta.get('curso', '')
                alumno = respuesta.get('alumno', '')
                
                if curso and alumno and alumno.strip():
                    for i in range(1, 16):
                        columna = f"m{i}"
                        valor = respuesta.get(columna, False)
                        valor_bool = convertir_a_booleano(valor)
                        key = f"{curso}_{alumno}_{i}"
                        estado[key] = valor_bool
                        contador_cargados += 1
        
        print(f"✅ Total estados cargados: {contador_cargados}")
        
        keys_to_remove = [key for key in session.keys() if key.startswith('estado_temp_')]
        for key in keys_to_remove:
            session.pop(key, None)
        
        for key, value in estado.items():
            session[f"estado_temp_{key}"] = value
        
        return render_template('panel.html',
                               cursos=cursos,
                               todas_materias=todas_materias,
                               materias_por_curso=materias_por_curso,
                               estado=estado,
                               usuario=usuario,
                               nombre_completo=session.get('nombre_completo', usuario))
    
    except Exception as e:
        print(f"❌ ERROR en panel: {e}")
        print(traceback.format_exc())
        return f"""
        <h1>Error al cargar el panel</h1>
        <p><strong>Error:</strong> {str(e)}</p>
        <pre>{traceback.format_exc()}</pre>
        <a href="/logout">Volver a intentar</a>
        """

@app.route('/guardar', methods=['POST'])
def guardar():
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return jsonify({"success": False, "error": mensaje}), 403
    
    try:
        data = request.json
        
        profesor = session.get('usuario', '').upper()
        curso = data.get('curso').strip()
        alumno = data.get('alumno').strip()
        materia = int(data.get('materia'))
        valor = data.get('valor')
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        valor_texto = booleano_a_texto(valor)
        
        # Guardar en sesión con clave ORIGINAL (sin normalizar)
        key = f"{curso}_{alumno}_{materia}"
        session[f"estado_temp_{key}"] = valor
        
        print(f"💾 Guardando: {profesor} - {curso} - {alumno} - M{materia} = {valor_texto}")
        
        # Buscar si ya existe registro para este curso+alumno usando NORMALIZACIÓN para evitar duplicados
        todas_filas = resp_sheet.get_all_values()
        
        num_fila = None
        curso_norm = normalizar_texto(curso)
        alumno_norm = normalizar_texto(alumno)
        
        for idx, fila in enumerate(todas_filas, start=1):
            if idx == 1:
                continue
            if len(fila) >= 3:
                fila_curso = fila[1].strip() if len(fila) > 1 else ""
                fila_alumno = fila[2].strip() if len(fila) > 2 else ""
                
                # Comparación NORMALIZADA para evitar duplicados
                if normalizar_texto(fila_curso) == curso_norm and normalizar_texto(fila_alumno) == alumno_norm:
                    num_fila = idx
                    break
        
        columna_materia = 4 + (materia - 1)
        
        if num_fila:
            # Actualizar la materia específica
            resp_sheet.update_cell(num_fila, columna_materia, valor_texto)
            resp_sheet.update_cell(num_fila, 1, profesor)
            
            if len(todas_filas[num_fila-1]) >= 19:
                resp_sheet.update_cell(num_fila, 19, fecha)
            else:
                valores_actuales = resp_sheet.row_values(num_fila)
                while len(valores_actuales) < 19:
                    valores_actuales.append('')
                valores_actuales[18] = fecha
                resp_sheet.update(f'A{num_fila}:S{num_fila}', [valores_actuales])
            
            print(f"   Actualizada fila {num_fila}")
        else:
            # Crear nueva fila con los nombres ORIGINALES (tal cual llegaron)
            nueva_fila = [profesor, curso, alumno]
            for i in range(15):
                nueva_fila.append("FALSE")
            nueva_fila.append(fecha)
            
            resp_sheet.append_row(nueva_fila)
            print(f"   Creada nueva fila para {alumno}")
            
            # Actualizar la materia específica en la fila recién creada
            todas_filas_nuevas = resp_sheet.get_all_values()
            for idx, fila in enumerate(todas_filas_nuevas, start=1):
                if idx == 1:
                    continue
                if len(fila) >= 3:
                    fila_curso = fila[1].strip() if len(fila) > 1 else ""
                    fila_alumno = fila[2].strip() if len(fila) > 2 else ""
                    
                    if normalizar_texto(fila_curso) == curso_norm and normalizar_texto(fila_alumno) == alumno_norm:
                        resp_sheet.update_cell(idx, columna_materia, valor_texto)
                        print(f"   Actualizada materia M{materia} en fila {idx}")
                        break
        
        return jsonify({"success": True, "mensaje": "Guardado correctamente"})
        
    except Exception as e:
        print(f"❌ Error en guardar: {e}")
        print(traceback.format_exc())
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/pdf')
def pdf():
    if 'usuario' not in session:
        return redirect('/')
    
    valido, mensaje = verificar_fecha_valida()
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    try:
        profesor = session['usuario']
        nombre_completo = session.get('nombre_completo', profesor)
        materias_por_curso = session.get('materias_por_curso', {})
        
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre and curso in session['cursos']:
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso] and nombre:
                    cursos[curso].append(nombre)
        
        materias_data = mat_sheet.get_all_records()
        todas_materias = {}
        for m in materias_data:
            try:
                id_materia = int(float(m.get('id', 0)))
                nombre_materia = str(m.get('nombre', '')).strip()
                if id_materia > 0 and nombre_materia and nombre_materia != '':
                    todas_materias[id_materia] = nombre_materia
            except:
                pass
        
        pdf_buffer = generar_reporte_pdf(profesor, cursos, todas_materias, materias_por_curso, solo_marcadas=True, nombre_completo=nombre_completo)
        
        incrementar_contador_descargas(profesor)
        
        nombre_limpio = nombre_completo.replace(' ', '_').replace('ñ', 'n').replace('Ñ', 'N')
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=f"reporte_{nombre_limpio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
    
    except Exception as e:
        print(f"❌ Error en PDF: {e}")
        print(traceback.format_exc())
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
            if len(fila) > 0 and fila[0].strip().upper() == session['usuario'].upper():
                registro = {
                    "fila_numero": i,
                    "profesor": fila[0] if len(fila) > 0 else "",
                    "curso": fila[1] if len(fila) > 1 else "",
                    "alumno": fila[2] if len(fila) > 2 else "",
                }
                for m in range(1, 16):
                    col_idx = 3 + (m - 1)
                    registro[f"m{m}"] = fila[col_idx] if len(fila) > col_idx else ""
                registro["fecha"] = fila[18] if len(fila) > 18 else ""
                resultado["mis_registros"].append(registro)
        
        return jsonify(resultado)
    
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/debug_estado')
def debug_estado():
    if 'usuario' not in session:
        return jsonify({"error": "No hay sesión"})
    
    estados = {}
    for key in session.keys():
        if key.startswith('estado_temp_'):
            estados[key] = session[key]
    
    return jsonify({
        "usuario": session['usuario'],
        "total_estados": len(estados),
        "estados": list(estados.keys())[:20],
        "ejemplo": dict(list(estados.items())[:5]) if estados else {}
    })

@app.route('/test_conexion')
def test_conexion():
    return jsonify({"success": True, "mensaje": "Servidor activo"})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
