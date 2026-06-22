import os
import glob
import time
import unicodedata
import pandas as pd
from playwright.sync_api import sync_playwright

# =================================================================
# 🔑 CONFIGURACIÓN DE CREDENCIALES AUTOMÁTICAS
# =================================================================
USUARIO_ICBF = "OSCARA.HOYOSF"
CONTRASEÑA_ICBF = "OSk4r-4n4-"


def normalizar_texto(texto):
    """Elimina tildes, diéresis, espacios extra y convierte a mayúsculas planas"""
    if not texto or str(texto).lower() == 'nan':
        return ""
    s = unicodedata.normalize('NFKD', str(texto))
    texto_plano = "".join([c for c in s if not unicodedata.combining(c)])
    return " ".join(texto_plano.upper().split())


def click_menu_por_texto(page, texto, n_indice=0):
    """Navega los menús a máxima velocidad usando auto-esperas nativas de Playwright"""
    for contexto in [page] + page.frames:
        try:
            coincidencias = contexto.get_by_text(texto, exact=False).filter(visible=True)
            if coincidencias.count() > n_indice:
                coincidencias.nth(n_indice).click(timeout=1500)
                return True
        except Exception:
            continue
    return False


def hacer_clic_por_alt(page, palabras_clave, nombre_boton):
    """Busca de forma nativa cualquier elemento por su atributo ALT o TITLE (Tu función ganadora de guardado)"""
    for contexto in [page] + page.frames:
        for keyword in palabras_clave:
            try:
                selectors = [
                    f"input[alt*='{keyword}' i]", 
                    f"img[alt*='{keyword}' i]",
                    f"[alt*='{keyword}' i]",
                    f"[title*='{keyword}' i]"
                ]
                for selector in selectors:
                    loc = contexto.locator(selector).first
                    if loc.is_visible():
                        loc.click(timeout=3000)
                        print(f"   🎯 ¡Clic exitoso automático en: {nombre_boton}!")
                        return True
            except Exception:
                continue
    return False


def iniciar_robot_excel_junio():
    todos_los_archivos = glob.glob("*.xlsx") + glob.glob("*.csv")
    archivos_excel = [f for f in todos_los_archivos if not os.path.basename(f).startswith("~$")]
    
    if not archivos_excel:
        print("❌ No se encontraron archivos de Excel o CSV en esta carpeta.")
        return

    print(f"📚 Se encontraron {len(archivos_excel)} archivos reales para procesar en lote.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        
        for numero, archivo_actual in enumerate(archivos_excel, 1):
            nombre_breve = os.path.basename(archivo_actual)
            print(f"\n==================================================================")
            print(f"📦 [{numero}/{len(archivos_excel)}] PROCESANDO: {nombre_breve}")
            print(f"==================================================================")
            
            context = browser.new_context()
            page = context.new_page()
            
            print("🔒 Abriendo portal...")
            page.goto("https://rubonline.icbf.gov.co/DefaultF.aspx")
            
            print("🔑 Logueando...")
            try:
                page.locator("input[type='text']").filter(visible=True).first.fill(USUARIO_ICBF)
                page.locator("input[type='password']").filter(visible=True).first.fill(CONTRASEÑA_ICBF)
                page.locator("input[type='submit'], button[type='submit'], :text('Ingresar')").filter(visible=True).first.click()
            except Exception:
                pass

            print("⚡ Navegando en ráfaga al módulo RAM...")
            time.sleep(2.5)  
            
            click_menu_por_texto(page, "Rub online", n_indice=1)
            time.sleep(0.5) 
            click_menu_por_texto(page, "Unidad", n_indice=0)
            time.sleep(0.5)
            click_menu_por_texto(page, "asistencia mensual", n_indice=0)

            print("\n⏳ ====== ACCIONES MANUALES REQUERIDAS EN LA WEB ======")
            print(f" UDS actual: '{nombre_breve}'")
            print(" Presiona Consultar -> Editar en la web.")
            
            input("👉 ENTER aquí apenas los cuadritos estén activos... ")
            
            try:
                if archivo_actual.endswith('.csv'):
                    df = pd.read_csv(archivo_actual, header=None)
                else:
                    df = pd.read_excel(archivo_actual, sheet_name='RAM', header=None)
            except Exception as e:
                print(f"⚠️ Error al leer Excel: {e}")
                context.close()
                continue

            print("\n🤖 Marcando a toda velocidad...")
            conteo_marcas = 0

            for indice, fila in df.iterrows():
                try:
                    primer_valor = str(fila.iloc[0]).strip()
                    if not primer_valor.isdigit():
                        continue 

                    p_nombre_ex = normalizar_texto(fila.iloc[3])
                    p_apellido_ex = normalizar_texto(fila.iloc[5])
                    
                    if not p_nombre_ex or not p_apellido_ex:
                        continue

                    fila_web = None

                    for fila_en_pantalla in page.locator("tr").all():
                        try:
                            texto_fila_limpio = normalizar_texto(fila_en_pantalla.inner_text())
                            if p_nombre_ex in texto_fila_limpio and p_apellido_ex in texto_fila_limpio:
                                if fila_en_pantalla.locator("input[type='checkbox']").count() < 40:
                                    fila_web = fila_en_pantalla
                                    break
                        except Exception:
                            continue

                    if not fila_web:
                        for frame in page.frames:
                            try:
                                for fila_en_pantalla in frame.locator("tr").all():
                                    texto_fila_limpio = normalizar_texto(fila_en_pantalla.inner_text())
                                    if p_nombre_ex in texto_fila_limpio and p_apellido_ex in texto_fila_limpio:
                                        if fila_en_pantalla.locator("input[type='checkbox']").count() < 40:
                                            fila_web = fila_en_pantalla
                                            break
                            except Exception:
                                continue
                            if fila_web:
                                break

                    if fila_web:
                        checkboxes = fila_web.locator("input[type='checkbox']")
                        marcas_nino = 0
                        
                        for dia in range(1, 31):
                            col_excel = 11 + dia  
                            if len(fila) > col_excel:
                                asistencia_dia = str(fila.iloc[col_excel]).strip().upper()
                                if asistencia_dia == "A":
                                    box_dia = checkboxes.nth(dia - 1)
                                    
                                    if box_dia.is_visible() and not box_dia.is_checked():
                                        box_dia.check()
                                        marcas_nino += 1
                                        conteo_marcas += 1
                                        
                        if marcas_nino > 0:
                            print(f"   ✅ {p_nombre_ex} {p_apellido_ex} -> Marcado rápido")
                    else:
                        print(f"   ⚠️ No localizado: {p_nombre_ex}")
                        
                except Exception:
                    continue
                    
            print(f"\n💾 Completado. Total clics: {conteo_marcas}")
            
            # --- 🔥 AQUÍ ENTRA TU JUGADA MAESTRA DE GUARDADO AUTOMÁTICO ---
            print("💾 Presionando el Disquete de Guardar automáticamente...")
            hacer_clic_por_alt(page, ["GUARDAR", "GRABAR"], "Disquete Guardar")
            
            input("\n👉 Revisa que haya guardado bien y presiona ENTER aquí para ir al siguiente archivo... ")
            print("------------------------------------------------------------------")
            context.close()  

        browser.close()
        print("\n🏆 ¡Lote completado a máxima velocidad y guardado automático, Oscar!")


if __name__ == "__main__":
    iniciar_robot_excel_junio()