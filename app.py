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

# ==================== FUNCIÓN DE NORMALIZACIÓN ====================
def normalizar_texto(texto):
    """Normaliza texto SOLO para comparación"""
    if not texto:
        return ""
    texto = str(texto).strip().upper()
    texto = ' '.join(texto.split())
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) 
                    if unicodedata.category(c) != 'Mn')
    return texto

def obtener_estados_desde_sheets(profesor):
    """Obtiene TODOS los estados desde Google Sheets para un profesor"""
    try:
        todas_respuestas = resp_sheet.get_all_records()
        estados = {}
        
        for respuesta in todas_respuestas:
            profesor_resp = respuesta.get('profesor', '')
            if profesor_resp and normalizar_texto(profesor_resp) == normalizar_texto(profesor):
                curso = respuesta.get('curso', '')
                alumno = respuesta.get('alumno', '')
                
                if curso and alumno:
                    for i in range(1, 21):
                        columna = f"m{i}"
                        valor = respuesta.get(columna, False)
                        valor_bool = convertir_a_booleano(valor)
                        key = f"{curso}_{alumno}_{i}"
                        estados[key] = valor_bool
        
        print(f"📥 Desde Sheets: {len(estados)} estados encontrados para {profesor}")
        return estados
    except Exception as e:
        print(f"❌ Error obteniendo estados desde Sheets: {e}")
        return {}

# ==================== INICIALIZACIÓN GOOGLE SHEETS ====================
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
    
    # Crear hoja CONFIG si no existe
    try:
        config_sheet = spreadsheet.worksheet("CONFIG")
    except:
        config_sheet = spreadsheet.add_worksheet("CONFIG", rows=10, cols=2)
        config_sheet.append_row(["clave", "valor"])
        config_sheet.append_row(["activo", "TRUE"])
        config_sheet.append_row(["fecha_inicio", "2026-01-01 00:00:00"])
        config_sheet.append_row(["fecha_fin", "2026-12-31 23:59:59"])
    
    # Crear hoja ESTADISTICAS si no existe
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
            if materia_id_int <= 0 or materia_id_int > 20:
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

def generar_reporte_pdf(profesor, cursos_data, todas_materias, materias_por_curso, estados_data, solo_marcadas=True, nombre_completo=""):
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
                        key = f"{curso_nombre}_{alumno}_{materia_id}"
                        if estados_data.get(key, False):
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
                key = f"{curso_nombre}_{alumno}_{materia_id}"
                estado_valor = estados_data.get(key, False)
                
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

# ==================== FUNCIONES DE ADMINISTRADOR ====================

def verificar_admin():
    """Verifica si el usuario actual es administrador"""
    admin_password = os.environ.get('ADMIN_PASSWORD', 'admin123')
    return session.get('admin_logged_in', False) and session.get('admin_password') == admin_password

# ==================== RUTAS DE PROFESOR ====================

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
        
        estados = obtener_estados_desde_sheets(usuario)
        
        keys_to_remove = [key for key in session.keys() if key.startswith('estado_temp_')]
        for key in keys_to_remove:
            session.pop(key, None)
        
        for key, value in estados.items():
            session[f"estado_temp_{key}"] = value
        
        return render_template('panel.html',
                               cursos=cursos,
                               todas_materias=todas_materias,
                               materias_por_curso=materias_por_curso,
                               estado=estados,
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
        
        key = f"{curso}_{alumno}_{materia}"
        session[f"estado_temp_{key}"] = valor
        
        print(f"💾 Guardando: {profesor} - {curso} - {alumno} - M{materia} = {valor_texto}")
        
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
                
                if normalizar_texto(fila_curso) == curso_norm and normalizar_texto(fila_alumno) == alumno_norm:
                    num_fila = idx
                    break
        
        columna_materia = 4 + (materia - 1)
        
        if num_fila:
            resp_sheet.update_cell(num_fila, columna_materia, valor_texto)
            resp_sheet.update_cell(num_fila, 1, profesor)
            
            if len(todas_filas[num_fila-1]) >= 24:
                resp_sheet.update_cell(num_fila, 24, fecha)
            else:
                valores_actuales = resp_sheet.row_values(num_fila)
                while len(valores_actuales) < 24:
                    valores_actuales.append('')
                valores_actuales[23] = fecha
                resp_sheet.update(f'A{num_fila}:X{num_fila}', [valores_actuales])
            
            print(f"   Actualizada fila {num_fila}")
        else:
            nueva_fila = [profesor, curso, alumno]
            for i in range(20):
                nueva_fila.append("FALSE")
            nueva_fila.append(fecha)
            
            resp_sheet.append_row(nueva_fila)
            print(f"   Creada nueva fila para {alumno}")
            
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
        
        estados = obtener_estados_desde_sheets(profesor)
        print(f"📊 PDF: {len(estados)} estados cargados desde Sheets para {profesor}")
        
        pdf_buffer = generar_reporte_pdf(profesor, cursos, todas_materias, materias_por_curso, estados, solo_marcadas=True, nombre_completo=nombre_completo)
        
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

# ==================== RUTAS DE ADMINISTRADOR ====================

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    """Login para administradores"""
    if request.method == 'POST':
        password = request.form.get('password', '')
        admin_password = os.environ.get('ADMIN_PASSWORD', 'admin123')
        
        if password == admin_password:
            session['admin_logged_in'] = True
            session['admin_password'] = password
            return redirect('/admin/dashboard')
        else:
            return render_template('admin_login.html', error="Contraseña incorrecta")
    
    return render_template('admin_login.html')

@app.route('/admin/dashboard')
def admin_dashboard():
    """Panel principal del administrador"""
    if not verificar_admin():
        return redirect('/admin/login')
    
    try:
        config_data = config_sheet.get_all_records()
        config = {}
        for row in config_data:
            config[row.get('clave', '')] = row.get('valor', '')
        
        stats = stats_sheet.get_all_records()
        
        profesores = prof_sheet.get_all_records()
        
        respuestas = resp_sheet.get_all_records()
        total_respuestas = len(respuestas)
        
        total_marcadas = 0
        for resp in respuestas:
            for i in range(1, 21):
                if convertir_a_booleano(resp.get(f'm{i}', False)):
                    total_marcadas += 1
        
        return render_template('admin_dashboard.html',
                             config=config,
                             stats=stats,
                             profesores=profesores,
                             total_profesores=len(profesores),
                             total_respuestas=total_respuestas,
                             total_marcadas=total_marcadas)
    
    except Exception as e:
        return f"Error: {e}"

@app.route('/admin/configurar_fechas', methods=['POST'])
def admin_configurar_fechas():
    """Configurar fechas del sistema"""
    if not verificar_admin():
        return jsonify({"success": False, "error": "No autorizado"}), 401
    
    try:
        activo = request.form.get('activo') == 'on'
        fecha_inicio = request.form.get('fecha_inicio')
        fecha_fin = request.form.get('fecha_fin')
        
        if fecha_inicio:
            fecha_inicio = fecha_inicio.replace('T', ' ') + ':00'
        if fecha_fin:
            fecha_fin = fecha_fin.replace('T', ' ') + ':00'
        
        config_sheet.update('B2', 'TRUE' if activo else 'FALSE')
        config_sheet.update('B3', fecha_inicio)
        config_sheet.update('B4', fecha_fin)
        
        return jsonify({"success": True, "mensaje": "Configuración actualizada"})
    
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/admin/profesores')
def admin_profesores():
    """Ver todos los profesores y sus asignaciones"""
    if not verificar_admin():
        return redirect('/admin/login')
    
    try:
        profesores = prof_sheet.get_all_records()
        materias = mat_sheet.get_all_records()
        materias_dict = {}
        for m in materias:
            try:
                id_materia = int(float(m.get('id', 0)))
                if id_materia > 0:
                    materias_dict[id_materia] = m.get('nombre', f'Materia {id_materia}')
            except:
                pass
        
        profesores_data = []
        for prof in profesores:
            profesor_info = {
                'usuario': prof.get('profesor', ''),
                'nombre_completo': prof.get('nombre_completo', ''),
                'materias': []
            }
            
            for i in range(1, 4):
                materia_id = prof.get(f'm{i}')
                cursos = prof.get(f'cursos_m{i}', '')
                if materia_id and str(materia_id).strip():
                    try:
                        materia_id_int = int(float(materia_id))
                        if materia_id_int in materias_dict:
                            profesor_info['materias'].append({
                                'materia': materias_dict[materia_id_int],
                                'cursos': cursos
                            })
                    except:
                        pass
            
            profesores_data.append(profesor_info)
        
        return render_template('admin_profesores.html',
                             profesores=profesores_data)
    
    except Exception as e:
        return f"Error: {e}"

@app.route('/admin/supervisar/<profesor>')
def admin_supervisar(profesor):
    """Supervisar las evaluaciones de un profesor específico"""
    if not verificar_admin():
        return redirect('/admin/login')
    
    try:
        todas_filas = prof_sheet.get_all_values()
        profesor_info = None
        encabezados = todas_filas[0] if todas_filas else []
        
        for fila in todas_filas[1:]:
            if len(fila) > 0 and fila[0].upper() == profesor.upper():
                profesor_info = {}
                for idx, header in enumerate(encabezados):
                    if idx < len(fila):
                        profesor_info[header] = fila[idx]
                    else:
                        profesor_info[header] = ''
                break
        
        if not profesor_info:
            return "Profesor no encontrado"
        
        cursos_profesor = []
        for i in range(1, 4):
            cursos_str = profesor_info.get(f'cursos_m{i}', '')
            if cursos_str:
                for c in str(cursos_str).split(','):
                    curso_limpio = c.strip()
                    if curso_limpio and curso_limpio not in cursos_profesor:
                        cursos_profesor.append(curso_limpio)
        
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre and curso in cursos_profesor:
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso]:
                    cursos[curso].append(nombre)
        
        materias_data = mat_sheet.get_all_records()
        todas_materias = {}
        for m in materias_data:
            try:
                id_materia = int(float(m.get('id', 0)))
                nombre_materia = str(m.get('nombre', '')).strip()
                if id_materia > 0 and nombre_materia:
                    todas_materias[id_materia] = nombre_materia
            except:
                pass
        
        profesor_dict = profesor_info
        materias_por_curso = {}
        for curso in cursos_profesor:
            materias_por_curso[curso] = []
        
        for i in range(1, 4):
            materia_id = profesor_dict.get(f'm{i}')
            cursos_str = profesor_dict.get(f'cursos_m{i}', '')
            
            if materia_id and str(materia_id).strip():
                try:
                    materia_id_int = int(float(materia_id))
                    if cursos_str:
                        for c in str(cursos_str).split(','):
                            curso_limpio = c.strip()
                            if curso_limpio in materias_por_curso:
                                if materia_id_int not in materias_por_curso[curso_limpio]:
                                    materias_por_curso[curso_limpio].append(materia_id_int)
                except:
                    pass
        
        estados = obtener_estados_desde_sheets(profesor)
        
        total_posibles = 0
        total_marcadas = 0
        
        for curso_nombre, alumnos in cursos.items():
            for alumno in alumnos:
                for materia_id in materias_por_curso.get(curso_nombre, []):
                    total_posibles += 1
                    key = f"{curso_nombre}_{alumno}_{materia_id}"
                    if estados.get(key, False):
                        total_marcadas += 1
        
        porcentaje = (total_marcadas / total_posibles * 100) if total_posibles > 0 else 0
        
        return render_template('admin_supervisar.html',
                             profesor=profesor,
                             profesor_info=profesor_info,
                             cursos=cursos,
                             todas_materias=todas_materias,
                             materias_por_curso=materias_por_curso,
                             estados=estados,
                             total_posibles=total_posibles,
                             total_marcadas=total_marcadas,
                             porcentaje=porcentaje)
    
    except Exception as e:
        return f"Error: {e}"

@app.route('/admin/estadisticas')
def admin_estadisticas():
    """Estadísticas detalladas del sistema"""
    if not verificar_admin():
        return redirect('/admin/login')
    
    try:
        stats = stats_sheet.get_all_records()
        
        profesores = prof_sheet.get_all_records()
        
        evaluaciones_por_profesor = {}
        respuestas = resp_sheet.get_all_records()
        
        for resp in respuestas:
            profesor = resp.get('profesor', '')
            if profesor:
                if profesor not in evaluaciones_por_profesor:
                    evaluaciones_por_profesor[profesor] = {'marcadas': 0, 'totales': 0}
                
                for i in range(1, 21):
                    evaluaciones_por_profesor[profesor]['totales'] += 1
                    if convertir_a_booleano(resp.get(f'm{i}', False)):
                        evaluaciones_por_profesor[profesor]['marcadas'] += 1
        
        profesores_nombres = []
        porcentajes = []
        descargas = []
        
        for prof in profesores:
            nombre = prof.get('profesor', '')
            if nombre:
                profesores_nombres.append(nombre)
                evaluaciones = evaluaciones_por_profesor.get(nombre, {'marcadas': 0, 'totales': 0})
                porcentaje = (evaluaciones['marcadas'] / evaluaciones['totales'] * 100) if evaluaciones['totales'] > 0 else 0
                porcentajes.append(porcentaje)
                
                descarga = 0
                for stat in stats:
                    if stat.get('profesor', '').upper() == nombre.upper():
                        descarga = int(stat.get('descargas_pdf', 0))
                        break
                descargas.append(descarga)
        
        return render_template('admin_estadisticas.html',
                             stats=stats,
                             profesores_nombres=profesores_nombres,
                             porcentajes=porcentajes,
                             descargas=descargas,
                             total_descargas=sum(descargas))
    
    except Exception as e:
        return f"Error: {e}"

@app.route('/admin/exportar_datos')
def admin_exportar_datos():
    """Exportar todos los datos a CSV"""
    if not verificar_admin():
        return redirect('/admin/login')
    
    try:
        import csv
        
        output = io.StringIO()
        
        respuestas = resp_sheet.get_all_records()
        
        if respuestas:
            writer = csv.DictWriter(output, fieldnames=respuestas[0].keys())
            writer.writeheader()
            writer.writerows(respuestas)
        
        output.seek(0)
        
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8-sig')),
            as_attachment=True,
            download_name=f"exportacion_respuestas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mimetype='text/csv'
        )
    
    except Exception as e:
        return f"Error: {e}"

@app.route('/admin/logout')
def admin_logout():
    """Cerrar sesión de administrador"""
    session.pop('admin_logged_in', None)
    session.pop('admin_password', None)
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
                for m in range(1, 21):
                    col_idx = 3 + (m - 1)
                    registro[f"m{m}"] = fila[col_idx] if len(fila) > col_idx else ""
                registro["fecha"] = fila[23] if len(fila) > 23 else ""
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
