# (TODO TU CÓDIGO ORIGINAL IGUAL HASTA /admin/panel)

@app.route('/admin/panel')
def admin_panel():
    if not session.get('admin'):
        return redirect('/admin')
    
    try:
        # ========== 1. LECTURA ÚNICA DE CADA HOJA ==========
        
        config = {'activo': True, 'fecha_inicio': '', 'fecha_fin': ''}
        try:
            valores_config = config_sheet.get_all_values()
            if len(valores_config) > 1:
                for fila in valores_config[1:]:
                    if len(fila) >= 2:
                        clave = str(fila[0]).strip() if fila[0] else ''
                        valor = str(fila[1]).strip() if len(fila) > 1 and fila[1] else ''
                        if clave == 'activo':
                            config['activo'] = convertir_a_booleano(valor)
                        elif clave == 'fecha_inicio':
                            config['fecha_inicio'] = valor.replace(' ', 'T') if valor else ''
                        elif clave == 'fecha_fin':
                            config['fecha_fin'] = valor.replace(' ', 'T') if valor else ''
        except Exception as e:
            print(f"⚠️ Error leyendo CONFIG: {e}")
        
        estudiantes_lista = []
        alumnos_por_curso = {}
        cursos_unicos = set()
        try:
            valores_est = est_sheet.get_all_values()
            if len(valores_est) > 1:
                for fila in valores_est[1:]:
                    if len(fila) >= 2:
                        curso = str(fila[0]).strip() if fila[0] else ''
                        nombre = str(fila[1]).strip() if len(fila) > 1 and fila[1] else ''
                        if curso and nombre:
                            estudiantes_lista.append({'curso': curso, 'nombre': nombre})
                            cursos_unicos.add(curso)
                            if curso not in alumnos_por_curso:
                                alumnos_por_curso[curso] = []
                            if nombre not in alumnos_por_curso[curso]:
                                alumnos_por_curso[curso].append(nombre)
        except Exception as e:
            print(f"⚠️ Error leyendo ESTUDIANTES: {e}")
        
        todas_materias = {}
        try:
            valores_mat = mat_sheet.get_all_values()
            if len(valores_mat) > 1:
                for fila in valores_mat[1:]:
                    if len(fila) >= 2 and fila[0] and fila[1]:
                        try:
                            id_mat = int(float(fila[0]))
                            todas_materias[id_mat] = str(fila[1]).strip()
                        except:
                            pass
        except Exception as e:
            print(f"⚠️ Error leyendo MATERIAS: {e}")
        
        profesores_data = []
        try:
            valores_prof = prof_sheet.get_all_values()
            if len(valores_prof) > 1:
                encabezados = [str(h).strip() if h else '' for h in valores_prof[0]]
                for fila in valores_prof[1:]:
                    prof_dict = {}
                    for i, header in enumerate(encabezados):
                        if header and i < len(fila):
                            prof_dict[header] = str(fila[i]).strip() if fila[i] else ''
                    usuario = prof_dict.get('usuario', '')
                    if usuario:
                        profesores_data.append(prof_dict)
        except Exception as e:
            print(f"⚠️ Error leyendo PROFESORES: {e}")
        
        todas_respuestas = []
        try:
            valores_resp = resp_sheet.get_all_values()
            if len(valores_resp) > 1:
                encabezados_resp = [str(h).strip() if h else '' for h in valores_resp[0]]
                for fila in valores_resp[1:]:
                    resp_dict = {}
                    for i, header in enumerate(encabezados_resp):
                        if header and i < len(fila):
                            resp_dict[header] = str(fila[i]).strip() if fila[i] else ''
                    if resp_dict.get('profesor', ''):
                        todas_respuestas.append(resp_dict)
        except Exception as e:
            print(f"⚠️ Error leyendo RESPUESTAS: {e}")
        
        stats_por_profesor = {}
        try:
            valores_stats = stats_sheet.get_all_values()
            if len(valores_stats) > 1:
                for fila in valores_stats[1:]:
                    if len(fila) >= 3:
                        prof = str(fila[0]).strip().upper() if fila[0] else ''
                        if prof:
                            stats_por_profesor[prof] = {
                                'descargas': str(fila[1]).strip() if len(fila) > 1 and fila[1] else '0',
                                'ultima': str(fila[2]).strip() if len(fila) > 2 and fila[2] else ''
                            }
        except Exception as e:
            print(f"⚠️ Error leyendo ESTADISTICAS: {e}")
        
        # ========== 2. PROCESAR EN MEMORIA ==========
        
        respuestas_por_profesor = {}
        total_evaluaciones = 0
        
        # 🔧 CORREGIDO: NORMALIZACIÓN AQUÍ
        for resp in todas_respuestas:
            profesor_resp = resp.get('profesor', '')
            
            if profesor_resp:
                profesor_norm = normalizar_texto(profesor_resp)
                
                if profesor_norm not in respuestas_por_profesor:
                    respuestas_por_profesor[profesor_norm] = {}
                
                curso = normalizar_texto(resp.get('curso', ''))
                alumno = normalizar_texto(resp.get('alumno', ''))
                
                if curso and alumno:
                    for i in range(1, 21):
                        columna = f"m{i}"
                        valor = resp.get(columna, '')
                        
                        if convertir_a_booleano(valor):
                            key = f"{curso}_{alumno}_{i}"
                            respuestas_por_profesor[profesor_norm][key] = True
                            total_evaluaciones += 1
        
        profesores = []
        for prof_dict in profesores_data:
            usuario = prof_dict.get('usuario', '')
            if usuario:
                cursos_set = set()
                materias_por_curso_prof = {}
                
                for i in range(1, 4):
                    cursos_str = prof_dict.get(f'cursos_m{i}', '')
                    materia_id_str = prof_dict.get(f'm{i}', '')
                    
                    if cursos_str:
                        for c in cursos_str.split(','):
                            curso_limpio = c.strip()
                            if curso_limpio:
                                cursos_set.add(curso_limpio)
                                if materia_id_str and materia_id_str.strip():
                                    if curso_limpio not in materias_por_curso_prof:
                                        materias_por_curso_prof[curso_limpio] = []
                                    try:
                                        id_mat = int(float(materia_id_str))
                                        if id_mat > 0 and id_mat <= 20:
                                            if id_mat not in materias_por_curso_prof[curso_limpio]:
                                                materias_por_curso_prof[curso_limpio].append(id_mat)
                                    except:
                                        pass
                
                cursos_lista = sorted(list(cursos_set))
                
                usuario_norm = normalizar_texto(usuario)
                estados_prof = respuestas_por_profesor.get(usuario_norm, {})
                
                cursos_detalle = []
                total_evaluaciones_prof = 0
                total_posibles_prof = 0
                
                for curso in cursos_lista:
                    alumnos_curso = alumnos_por_curso.get(curso, [])
                    materias_curso = materias_por_curso_prof.get(curso, [])
                    
                    # 🔧 CORREGIDO: NORMALIZACIÓN AQUÍ
                    curso_norm = normalizar_texto(curso)
                    
                    evaluadas = 0
                    posibles = len(alumnos_curso) * len(materias_curso)
                    
                    for alumno in alumnos_curso:
                        alumno_norm = normalizar_texto(alumno)
                        
                        for materia in materias_curso:
                            key = f"{curso_norm}_{alumno_norm}_{materia}"
                            if estados_prof.get(key, False):
                                evaluadas += 1
                    
                    total_evaluaciones_prof += evaluadas
                    total_posibles_prof += posibles
                    
                    porcentaje = (evaluadas / posibles * 100) if posibles > 0 else 0
                    
                    cursos_detalle.append({
                        'nombre': curso,
                        'alumnos': len(alumnos_curso),
                        'materias': len(materias_curso),
                        'evaluadas': evaluadas,
                        'posibles': posibles,
                        'porcentaje': round(porcentaje, 1)
                    })
                
                stats = stats_por_profesor.get(usuario.upper(), {})
                descargas = 0
                try:
                    descargas = int(stats.get('descargas', '0'))
                except:
                    pass
                
                porcentaje_general = (total_evaluaciones_prof / total_posibles_prof * 100) if total_posibles_prof > 0 else 0
                
                profesores.append({
                    'usuario': usuario,
                    'nombre_completo': prof_dict.get('nombre_completo', usuario),
                    'cursos': cursos_lista,
                    'cursos_detalle': cursos_detalle,
                    'descargas': descargas,
                    'ultima_descarga': stats.get('ultima', ''),
                    'total_evaluaciones': total_evaluaciones_prof,
                    'total_posibles': total_posibles_prof,
                    'porcentaje_general': round(porcentaje_general, 1)
                })
        
        cursos_materias = {}
        for curso in sorted(cursos_unicos):
            cursos_materias[curso] = []
            for id_mat, nombre_mat in sorted(todas_materias.items()):
                cursos_materias[curso].append({'id': id_mat, 'nombre': nombre_mat})
        
        return render_template('admin_panel.html',
                             config=config,
                             profesores=profesores,
                             cursos_materias=cursos_materias,
                             total_profesores=len(profesores),
                             total_cursos=len(cursos_unicos),
                             total_alumnos=len(estudiantes_lista),
                             total_evaluaciones=total_evaluaciones)
    
    except Exception as e:
        print(f"❌ Error en panel admin: {e}")
        print(traceback.format_exc())
        return render_template('admin_panel.html', error=str(e))
