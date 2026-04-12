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
        
        print(f"🔍 Buscando estados para profesor: '{profesor}'")
        
        for respuesta in todas_respuestas:
            profesor_resp = respuesta.get('profesor', '')
            profesor_resp_norm = normalizar_texto(profesor_resp)
            profesor_norm = normalizar_texto(profesor)
            
            if profesor_resp_norm == profesor_norm:
                curso = respuesta.get('curso', '')
                alumno = respuesta.get('alumno', '')
                
                if curso and alumno:
                    for i in range(1, 21):
                        columna = f"m{i}"
                        valor = respuesta.get(columna, False)
                        valor_bool = convertir_a_booleano(valor)
                        key = f"{curso}_{alumno}_{i}"
                        estados[key] = valor_bool
                        if valor_bool:
                            print(f"   ✅ Marcada: {curso} - {alumno} - M{i}")
        
        print(f"📥 Desde Sheets: {len(estados)} estados encontrados para {profesor}")
        return estados
    except Exception as e:
        print(f"❌ Error obteniendo estados desde Sheets: {e}")
        traceback.print_exc()
        return {}

# ====================================================================

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

def verificar_fecha_valida(ignorar_admin=False):
    """Verifica si el sistema está activo (ignorar_admin=True para rutas de admin)"""
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
        
        if not ignorar_admin:
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
    """Genera reporte PDF usando estados PASADOS como parámetro (desde Sheets)"""
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

# ==================== FUNCIONES ADMINISTRATIVAS ====================

def verificar_admin():
    """Verifica si el usuario actual es administrador"""
    return session.get('admin', False)

def obtener_todos_profesores():
    """Obtiene lista de todos los profesores con sus cursos"""
    try:
        todas_filas = prof_sheet.get_all_values()
        profesores = []
        
        if len(todas_filas) < 2:
            return []
        
        for fila in todas_filas[1:]:
            if len(fila) > 0 and fila[0].strip():
                profesor = {
                    'usuario': fila[0].strip(),
                    'nombre_completo': fila[1].strip() if len(fila) > 1 else '',
                    'cursos': []
                }
                
                for i in range(1, 4):
                    idx_curso = 3 + (i-1)*2
                    if len(fila) > idx_curso:
                        cursos_str = fila[idx_curso].strip() if fila[idx_curso] else ''
                        if cursos_str:
                            for c in cursos_str.split(','):
                                curso_limpio = c.strip()
                                if curso_limpio and curso_limpio not in profesor['cursos']:
                                    profesor['cursos'].append(curso_limpio)
                
                profesor['cursos'].sort()
                profesores.append(profesor)
        
        return profesores
    except Exception as e:
        print(f"Error obteniendo profesores: {e}")
        return []

def obtener_estadisticas_profesores():
    """Obtiene estadísticas de descargas por profesor"""
    try:
        return stats_sheet.get_all_records()
    except Exception as e:
        print(f"Error obteniendo estadísticas: {e}")
        return []

def actualizar_config_sistema(activo, fecha_inicio, fecha_fin):
    """Actualiza la configuración del sistema"""
    try:
        config_data = config_sheet.get_all_records()
        
        for idx, row in enumerate(config_data, start=2):
            clave = row.get('clave', '')
            if clave == 'activo':
                config_sheet.update_cell(idx, 2, 'TRUE' if activo else 'FALSE')
            elif clave == 'fecha_inicio':
                config_sheet.update_cell(idx, 2, fecha_inicio if fecha_inicio else '')
            elif clave == 'fecha_fin':
                config_sheet.update_cell(idx, 2, fecha_fin if fecha_fin else '')
        
        return True, "Configuración actualizada correctamente"
    except Exception as e:
        return False, str(e)

def obtener_config_actual():
    """Obtiene la configuración actual del sistema"""
    try:
        config_data = config_sheet.get_all_records()
        config = {
            'activo': True,
            'fecha_inicio': '',
            'fecha_fin': ''
        }
        
        for row in config_data:
            clave = row.get('clave', '')
            valor = row.get('valor', '')
            if clave == 'activo':
                config['activo'] = convertir_a_booleano(valor)
            elif clave == 'fecha_inicio':
                config['fecha_inicio'] = valor
            elif clave == 'fecha_fin':
                config['fecha_fin'] = valor
        
        return config
    except Exception as e:
        print(f"Error obteniendo config: {e}")
        return {'activo': True, 'fecha_inicio': '', 'fecha_fin': ''}

def obtener_datos_profesor(profesor_usuario):
    """Obtiene datos completos de un profesor específico"""
    try:
        todas_filas = prof_sheet.get_all_values()
        profesor_data = None
        
        if len(todas_filas) < 2:
            return None
        
        for fila in todas_filas[1:]:
            if len(fila) > 0 and fila[0].strip().upper() == profesor_usuario.upper():
                profesor_data = {
                    'usuario': fila[0].strip(),
                    'nombre_completo': fila[1].strip() if len(fila) > 1 else '',
                    'cursos': []
                }
                
                # Mapeo correcto de columnas según estructura de PROFESORES
                # Col 0: usuario, Col 1: nombre_completo
                # Col 2: m1, Col 3: cursos_m1
                # Col 4: m2, Col 5: cursos_m2
                # Col 6: m3, Col 7: cursos_m3
                for i in range(1, 4):
                    idx_materia = 2 + (i-1)*2  # 2, 4, 6
                    idx_cursos = 3 + (i-1)*2   # 3, 5, 7
                    
                    materia = fila[idx_materia].strip() if len(fila) > idx_materia else ''
                    cursos_str = fila[idx_cursos].strip() if len(fila) > idx_cursos else ''
                    
                    profesor_data[f'm{i}'] = materia
                    profesor_data[f'cursos_m{i}'] = cursos_str
                    
                    if cursos_str:
                        for c in cursos_str.split(','):
                            curso_limpio = c.strip()
                            if curso_limpio and curso_limpio not in profesor_data['cursos']:
                                profesor_data['cursos'].append(curso_limpio)
                
                profesor_data['cursos'].sort()
                break
        
        if not profesor_data:
            return None
        
        todas_respuestas = resp_sheet.get_all_records()
        evaluaciones = []
        
        for respuesta in todas_respuestas:
            if respuesta.get('profesor', '').upper() == profesor_usuario.upper():
                evaluacion = {
                    'curso': respuesta.get('curso', ''),
                    'alumno': respuesta.get('alumno', ''),
                    'fecha': respuesta.get('fecha', ''),
                }
                materias_marcadas = 0
                for i in range(1, 21):
                    if convertir_a_booleano(respuesta.get(f'm{i}', False)):
                        materias_marcadas += 1
                evaluacion['total_materias'] = materias_marcadas
                evaluaciones.append(evaluacion)
        
        profesor_data['evaluaciones'] = evaluaciones
        profesor_data['total_evaluaciones'] = len(evaluaciones)
        
        return profesor_data
        
    except Exception as e:
        print(f"Error obteniendo datos del profesor: {e}")
        traceback.print_exc()
        return None

# ==================== RUTAS DE ADMINISTRACIÓN ====================

@app.route('/admin', methods=['GET', 'POST'])
def admin_login():
    """Login para administradores (solo contraseña)"""
    valido, mensaje = verificar_fecha_valida(ignorar_admin=True)
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    if request.method == 'POST':
        password = request.form.get('password', '').strip()
        ADMIN_PASS = os.environ.get('ADMIN_PASS', 'admin123')
        
        if password == ADMIN_PASS:
            session['admin'] = True
            return redirect('/admin/panel')
        else:
            return render_template('admin_login.html', error="Contraseña incorrecta")
    
    return render_template('admin_login.html')

@app.route('/admin/panel')
def admin_panel():
    """Panel principal de administración"""
    if not verificar_admin():
        return redirect('/admin')
    
    valido, mensaje = verificar_fecha_valida(ignorar_admin=True)
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    try:
        profesores = obtener_todos_profesores()
        estadisticas = obtener_estadisticas_profesores()
        config = obtener_config_actual()
        
        todas_respuestas = resp_sheet.get_all_records()
        total_evaluaciones = len(todas_respuestas)
        
        total_marcadas = 0
        for respuesta in todas_respuestas:
            for i in range(1, 21):
                if convertir_a_booleano(respuesta.get(f'm{i}', False)):
                    total_marcadas += 1
        
        for profesor in profesores:
            profesor['descargas'] = 0
            for stat in estadisticas:
                if stat.get('profesor', '').upper() == profesor['usuario'].upper():
                    profesor['descargas'] = int(stat.get('descargas_pdf', 0))
                    break
        
        return render_template('admin_panel.html',
                             profesores=profesores,
                             config=config,
                             total_profesores=len(profesores),
                             total_evaluaciones=total_evaluaciones,
                             total_marcadas=total_marcadas)
    
    except Exception as e:
        print(f"Error en admin_panel: {e}")
        print(traceback.format_exc())
        return render_template('admin_error.html', mensaje=str(e))

@app.route('/admin/config', methods=['GET', 'POST'])
def admin_config():
    """Configuración del sistema (activo/fechas)"""
    if not verificar_admin():
        return redirect('/admin')
    
    valido, mensaje = verificar_fecha_valida(ignorar_admin=True)
    if not valido:
        return render_template('expirado.html', mensaje=mensaje)
    
    if request.method == 'POST':
        try:
            activo = request.form.get('activo') == 'on'
            fecha_inicio = request.form.get('fecha_inicio', '')
            fecha_fin = request.form.get('fecha_fin', '')
            
            if fecha_inicio:
                try:
                    datetime.strptime(fecha_inicio, "%Y-%m-%dT%H:%M")
                    fecha_inicio = fecha_inicio.replace('T', ' ') + ':00'
                except:
                    return render_template('admin_config.html',
                                         config=obtener_config_actual(),
                                         error="Formato de fecha inicio inválido")
            
            if fecha_fin:
                try:
                    datetime.strptime(fecha_fin, "%Y-%m-%dT%H:%M")
                    fecha_fin = fecha_fin.replace('T', ' ') + ':00'
                except:
                    return render_template('admin_config.html',
                                         config=obtener_config_actual(),
                                         error="Formato de fecha fin inválido")
            
            exito, mensaje = actualizar_config_sistema(activo, fecha_inicio, fecha_fin)
            
            if exito:
                return render_template('admin_config.html',
                                     config=obtener_config_actual(),
                                     mensaje=mensaje)
            else:
                return render_template('admin_config.html',
                                     config=obtener_config_actual(),
                                     error=mensaje)
        
        except Exception as e:
            return render_template('admin_config.html',
                                 config=obtener_config_actual(),
                                 error=str(e))
    
    config = obtener_config_actual()
    if config['fecha_inicio'] and ' ' in config['fecha_inicio']:
        config['fecha_inicio_display'] = config['fecha_inicio'].replace(' ', 'T')[:-3]
    else:
        config['fecha_inicio_display'] = ''
    
    if config['fecha_fin'] and ' ' in config['fecha_fin']:
        config['fecha_fin_display'] = config['fecha_fin'].replace(' ', 'T')[:-3]
    else:
        config['fecha_fin_display'] = ''
    
    return render_template('admin_config.html', config=config)

@app.route('/admin/profesor/<usuario>')
def admin_profesor_detalle(usuario):
    """Detalle de un profesor específico con materias marcadas por curso"""
    if not verificar_admin():
        return redirect('/admin')
    
    try:
        profesor_data = obtener_datos_profesor(usuario)
        
        if not profesor_data:
            return render_template('admin_error.html', mensaje=f"Profesor '{usuario}' no encontrado")
        
        # Obtener estudiantes para los cursos del profesor
        estudiantes = est_sheet.get_all_records()
        cursos_con_estudiantes = {}
        
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre and curso in profesor_data.get('cursos', []):
                if curso not in cursos_con_estudiantes:
                    cursos_con_estudiantes[curso] = []
                if nombre not in cursos_con_estudiantes[curso]:
                    cursos_con_estudiantes[curso].append(nombre)
        
        # Ordenar estudiantes alfabéticamente
        for curso in cursos_con_estudiantes:
            cursos_con_estudiantes[curso].sort()
        
        # Obtener materias
        materias_data = mat_sheet.get_all_records()
        todas_materias = {}
        for m in materias_data:
            try:
                id_materia = int(float(m.get('id', 0)))
                nombre_materia = str(m.get('nombre', '')).strip()
                if id_materia > 0 and nombre_materia:
                    todas_materias[str(id_materia)] = nombre_materia
            except:
                pass
        
        # Obtener materias por curso para este profesor
        todas_filas = prof_sheet.get_all_values()
        profesor_dict = {}
        if len(todas_filas) >= 2:
            for fila in todas_filas[1:]:
                if len(fila) > 0 and fila[0].strip().upper() == usuario.upper():
                    encabezados = todas_filas[0]
                    for idx, header in enumerate(encabezados):
                        if header and idx < len(fila):
                            profesor_dict[header] = fila[idx]
                    break
        
        materias_por_curso = obtener_materias_por_curso(profesor_dict, profesor_data.get('cursos', []))
        
        # Obtener estados desde Sheets para este profesor
        estados = obtener_estados_desde_sheets(usuario)
        
        # Construir matriz de evaluaciones por curso, alumno y materia
        evaluaciones_detalle = {}
        estadisticas_curso = {}
        
        for curso, alumnos in cursos_con_estudiantes.items():
            evaluaciones_detalle[curso] = {}
            materias_ids = materias_por_curso.get(curso, [])
            
            # Inicializar estadísticas
            estadisticas_curso[curso] = {
                'total_alumnos': len(alumnos),
                'total_materias': len(materias_ids),
                'total_posibles': len(alumnos) * len(materias_ids),
                'total_marcadas': 0,
                'materias_marcadas_por_alumno': {}
            }
            
            for alumno in alumnos:
                evaluaciones_detalle[curso][alumno] = {}
                materias_marcadas_alumno = 0
                
                for materia_id in materias_ids:
                    key = f"{curso}_{alumno}_{materia_id}"
                    esta_marcada = estados.get(key, False)
                    evaluaciones_detalle[curso][alumno][str(materia_id)] = esta_marcada
                    
                    if esta_marcada:
                        estadisticas_curso[curso]['total_marcadas'] += 1
                        materias_marcadas_alumno += 1
                
                estadisticas_curso[curso]['materias_marcadas_por_alumno'][alumno] = materias_marcadas_alumno
            
            # Calcular porcentaje
            if estadisticas_curso[curso]['total_posibles'] > 0:
                estadisticas_curso[curso]['porcentaje'] = round(
                    (estadisticas_curso[curso]['total_marcadas'] / 
                     estadisticas_curso[curso]['total_posibles']) * 100, 1
                )
            else:
                estadisticas_curso[curso]['porcentaje'] = 0
        
        # Obtener estadísticas de descargas
        try:
            stats = stats_sheet.get_all_records()
            descargas_profesor = 0
            for stat in stats:
                if stat.get('profesor', '').upper() == usuario.upper():
                    descargas_profesor = int(stat.get('descargas_pdf', 0))
                    break
        except:
            descargas_profesor = 0
        
        return render_template('admin_profesor_detalle.html', 
                             profesor=profesor_data,
                             cursos_con_estudiantes=cursos_con_estudiantes,
                             todas_materias=todas_materias,
                             materias_por_curso=materias_por_curso,
                             evaluaciones_detalle=evaluaciones_detalle,
                             estadisticas_curso=estadisticas_curso,
                             descargas_profesor=descargas_profesor,
                             total_estados_encontrados=len(estados))
    
    except Exception as e:
        print(f"Error en admin_profesor_detalle: {e}")
        traceback.print_exc()
        return render_template('admin_error.html', mensaje=str(e))

@app.route('/admin/reporte_individual/<usuario>')
def admin_reporte_individual(usuario):
    """Genera reporte PDF individual para un profesor específico"""
    if not verificar_admin():
        return redirect('/admin')
    
    try:
        profesor_data = obtener_datos_profesor(usuario)
        
        if not profesor_data:
            return "Profesor no encontrado", 404
        
        # Obtener estudiantes para los cursos del profesor
        estudiantes = est_sheet.get_all_records()
        cursos = {}
        
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre and curso in profesor_data.get('cursos', []):
                if curso not in cursos:
                    cursos[curso] = []
                if nombre not in cursos[curso]:
                    cursos[curso].append(nombre)
        
        # Ordenar estudiantes alfabéticamente
        for curso in cursos:
            cursos[curso].sort()
        
        # Obtener materias
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
        
        # Obtener datos del profesor de la hoja PROFESORES
        todas_filas = prof_sheet.get_all_values()
        profesor_dict = {}
        if len(todas_filas) >= 2:
            encabezados = todas_filas[0]
            for fila in todas_filas[1:]:
                if len(fila) > 0 and fila[0].strip().upper() == usuario.upper():
                    for idx, header in enumerate(encabezados):
                        if header and idx < len(fila):
                            profesor_dict[header] = fila[idx]
                    break
        
        materias_por_curso = obtener_materias_por_curso(profesor_dict, profesor_data.get('cursos', []))
        
        # Obtener estados desde Sheets
        estados = obtener_estados_desde_sheets(usuario)
        
        # solo_marcadas=False para mostrar TODAS las materias
        pdf_buffer = generar_reporte_pdf(usuario, cursos, todas_materias, materias_por_curso, estados, solo_marcadas=False, nombre_completo=profesor_data.get('nombre_completo', usuario))
        
        nombre_limpio = profesor_data.get('nombre_completo', usuario).replace(' ', '_')
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=f"reporte_individual_{nombre_limpio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
    
    except Exception as e:
        print(f"Error generando reporte individual: {e}")
        print(traceback.format_exc())
        return f"Error al generar PDF: {e}"

@app.route('/admin/reporte_grupal')
def admin_reporte_grupal():
    """Genera reporte PDF grupal de todos los profesores"""
    if not verificar_admin():
        return redirect('/admin')
    
    try:
        profesores = obtener_todos_profesores()
        
        if not profesores:
            return "No hay profesores registrados", 404
        
        # Obtener estudiantes una sola vez
        estudiantes = est_sheet.get_all_records()
        estudiantes_por_curso = {}
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre:
                if curso not in estudiantes_por_curso:
                    estudiantes_por_curso[curso] = []
                if nombre not in estudiantes_por_curso[curso]:
                    estudiantes_por_curso[curso].append(nombre)
        
        # Obtener materias
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
        
        # Generar PDF grupal
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
        
        profesor_style = ParagraphStyle(
            'ProfesorStyle',
            parent=styles['Heading3'],
            fontSize=14,
            spaceAfter=10,
            spaceBefore=15,
            textColor=colors.HexColor('#4a6cf7')
        )
        
        elementos = []
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        titulo = Paragraph(f"<b>REPORTE GENERAL DE EVALUACIONES</b>", titulo_style)
        elementos.append(titulo)
        subtitulo = Paragraph(f"Reporte de todos los profesores<br/>Fecha: {fecha_actual}", subtitulo_style)
        elementos.append(subtitulo)
        elementos.append(Spacer(1, 20))
        
        for profesor in profesores:
            elementos.append(Paragraph(f"<b>Profesor: {profesor['nombre_completo']} ({profesor['usuario']})</b>", profesor_style))
            elementos.append(Spacer(1, 10))
            
            # Obtener estados para este profesor
            estados = obtener_estados_desde_sheets(profesor['usuario'])
            
            # Obtener cursos del profesor con estudiantes
            cursos = {}
            for curso in profesor['cursos']:
                if curso in estudiantes_por_curso:
                    cursos[curso] = estudiantes_por_curso[curso]
            
            # Obtener materias por curso para este profesor
            todas_filas = prof_sheet.get_all_values()
            profesor_dict = {}
            if len(todas_filas) >= 2:
                for fila in todas_filas[1:]:
                    if len(fila) > 0 and fila[0].strip().upper() == profesor['usuario'].upper():
                        encabezados = todas_filas[0]
                        for idx, header in enumerate(encabezados):
                            if header and idx < len(fila):
                                profesor_dict[header] = fila[idx]
                        break
            
            materias_por_curso = obtener_materias_por_curso(profesor_dict, profesor['cursos'])
            
            for curso_nombre, alumnos in cursos.items():
                elementos.append(Paragraph(f"<b>Curso: {curso_nombre}</b>", styles['Heading4']))
                
                materias_curso_ids = materias_por_curso.get(curso_nombre, [])
                if not materias_curso_ids:
                    elementos.append(Paragraph("<i>No hay materias asignadas</i>", styles['Italic']))
                    elementos.append(Spacer(1, 10))
                    continue
                
                total_marcadas_curso = 0
                total_posibles_curso = 0
                
                for alumno in alumnos:
                    for materia_id in materias_curso_ids:
                        key = f"{curso_nombre}_{alumno}_{materia_id}"
                        if estados.get(key, False):
                            total_marcadas_curso += 1
                        total_posibles_curso += 1
                
                if total_posibles_curso > 0:
                    porcentaje = (total_marcadas_curso / total_posibles_curso) * 100
                    elementos.append(Paragraph(
                        f"<i>{len(alumnos)} alumnos - {total_marcadas_curso} evaluaciones de {total_posibles_curso} posibles ({porcentaje:.1f}%)</i>",
                        styles['Italic']
                    ))
                
                elementos.append(Spacer(1, 10))
            
            elementos.append(Spacer(1, 20))
        
        doc.build(elementos)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"reporte_grupal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
    
    except Exception as e:
        print(f"Error generando reporte grupal: {e}")
        print(traceback.format_exc())
        return f"Error al generar PDF: {e}"

@app.route('/admin/logout')
def admin_logout():
    """Cierre de sesión administrador"""
    session.pop('admin', None)
    return redirect('/admin')

# ==================== RUTAS ORIGINALES (SIN MODIFICAR) ====================

@app.route('/', methods=['GET', 'POST'])
def login():
    valido, mensaje = verificar_fecha_valida(ignorar_admin=False)
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
    
    valido, mensaje = verificar_fecha_valida(ignorar_admin=False)
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
    valido, mensaje = verificar_fecha_valida(ignorar_admin=False)
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
    
    valido, mensaje = verificar_fecha_valida(ignorar_admin=False)
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

# ==================== RUTAS DE DIAGNÓSTICO PARA ADMIN ====================

@app.route('/admin/ver_estados/<usuario>')
def admin_ver_estados(usuario):
    """Ruta de diagnóstico para ver qué estados tiene un profesor"""
    if not verificar_admin():
        return redirect('/admin')
    
    try:
        estados = obtener_estados_desde_sheets(usuario)
        
        # Obtener datos del profesor
        profesor_data = obtener_datos_profesor(usuario)
        
        # Obtener estudiantes
        estudiantes = est_sheet.get_all_records()
        cursos_con_estudiantes = {}
        for est in estudiantes:
            curso = str(est.get('curso', '')).strip()
            nombre = str(est.get('nombre', '')).strip()
            if curso and nombre and profesor_data and curso in profesor_data.get('cursos', []):
                if curso not in cursos_con_estudiantes:
                    cursos_con_estudiantes[curso] = []
                if nombre not in cursos_con_estudiantes[curso]:
                    cursos_con_estudiantes[curso].append(nombre)
        
        # Mostrar los estados encontrados
        return jsonify({
            "profesor": usuario,
            "total_estados_encontrados": len(estados),
            "estados": {k: v for k, v in list(estados.items())[:50]},
            "cursos": cursos_con_estudiantes,
            "claves_ejemplo": list(estados.keys())[:10]
        })
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
