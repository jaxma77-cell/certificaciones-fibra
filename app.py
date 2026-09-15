# --- FUNCIÓN BORRAR EN GITHUB (CORREGIDA CON MAPEO DE NOMBRES) ---
def borrar_en_github(tipo, valor, nombre_proyecto_app):
    """
    Borra filas del archivo datos_bot.json en GitHub
    Detecta automáticamente el nombre del proyecto en GitHub
    """
    if not GITHUB_TOKEN:
        return False, "❌ No hay GITHUB_TOKEN configurado"
    
    try:
        # 1. Leer archivo de GitHub
        url = f"https://api.github.com/repos/{GITHUB_USUARIO}/{GITHUB_REPO}/contents/{GITHUB_ARCHIVO_BOT}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        r = requests.get(url, headers=headers)
        
        if r.status_code != 200:
            return False, f"❌ Error al leer GitHub: {r.status_code}"
        
        contenido = base64.b64decode(r.json()["content"]).decode("utf-8")
        sha = r.json()["sha"]
        datos = json.loads(contenido)
        
        # 2. Buscar el proyecto en GitHub (con mapeo inteligente)
        proyectos_github = list(datos.keys())
        
        # Mapeo de nombres: app -> github
        mapeo = {
            "FIBRAMOL JUNIO Y JULIO": "FIBRAMOL",
            "Santomera Mayo 2024": "SANTOMERA",
            "FIBRAMOL": "FIBRAMOL",
            "SANTOMERA": "SANTOMERA"
        }
        
        # Buscar coincidencia
        nombre_github = None
        
        # Opción A: Coincidencia exacta
        if nombre_proyecto_app in datos:
            nombre_github = nombre_proyecto_app
        # Opción B: Usar mapeo
        elif nombre_proyecto_app in mapeo and mapeo[nombre_proyecto_app] in datos:
            nombre_github = mapeo[nombre_proyecto_app]
        # Opción C: Buscar por palabra clave
        else:
            for proj_github in proyectos_github:
                if proj_github.upper() in nombre_proyecto_app.upper() or nombre_proyecto_app.upper() in proj_github.upper():
                    nombre_github = proj_github
                    break
        
        if not nombre_github:
            return False, f"❌ No se encontró el proyecto en GitHub. Proyectos disponibles: {proyectos_github}"
        
        # 3. Filtrar filas
        filas_originales = datos[nombre_github].get("filas", [])
        
        if tipo == 'dia':
            filas_filtradas = [f for f in filas_originales if f.get('Dia') != valor]
            campo = 'Día'
        elif tipo == 'ubicacion':
            filas_filtradas = [f for f in filas_originales if f.get('Nombre') != valor]
            campo = 'Ubicación'
        else:
            return False, "❌ Tipo no válido"
        
        num_borradas = len(filas_originales) - len(filas_filtradas)
        
        if num_borradas == 0:
            return False, f"️ No se encontraron filas con {campo}='{valor}' en GitHub (proyecto: {nombre_github})"
        
        # 4. Guardar cambios
        datos[nombre_github]["filas"] = filas_filtradas
        
        contenido_nuevo = json.dumps(datos, indent=2, ensure_ascii=False)
        contenido_b64 = base64.b64encode(contenido_nuevo.encode("utf-8")).decode("utf-8")
        
        payload = {
            "message": f"Borrado {num_borradas} filas por {campo}: {valor} (desde app)",
            "content": contenido_b64,
            "sha": sha
        }
        
        r = requests.put(url, headers=headers, json=payload)
        
        if r.status_code in (200, 201):
            return True, f"✅ Borradas {num_borradas} filas de GitHub (proyecto: {nombre_github})"
        else:
            return False, f"❌ Error al guardar: {r.status_code} - {r.text}"
            
    except Exception as e:
        return False, f"❌ Error: {str(e)}"
        
