import os
import shutil
from datetime import datetime
import time
import xlwings as xw
import win32com.client
from pathlib import Path

class UnoBiableUpdater:
    def __init__(self):
        self.BASE_DIR =  r"C:\Users\aprsistemas\Desktop\trabajo\automatizacion_resportes"
        self.DATA_DIR = os.path.join(self.BASE_DIR, "data")
        self.INPUT_FILENAME = "Carex COL Reporte Vendedor.xlsx"
        self.INPUT_PATH = os.path.join(self.DATA_DIR, self.INPUT_FILENAME)
        self.BACKUP_DIR = os.path.join(self.DATA_DIR, "backups")
        os.makedirs(self.BACKUP_DIR, exist_ok=True)
        
        # Configuración para timeouts
        self.MAX_REFRESH_TIME = 20  # 5 minutos máximo para actualizar
        self.RETRY_ATTEMPTS = 2
    
    def verificar_archivo_disponible(self):
        """Verifica que el archivo no esté siendo usado por otro proceso."""
        try:
            # Intentar abrir el archivo en modo exclusivo
            with open(self.INPUT_PATH, 'r+b') as f:
                pass
            return True
        except (IOError, PermissionError):
            return False
    
    def hacer_backup(self):
        """Crea una copia de seguridad del archivo antes de modificarlo."""
        fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(self.BACKUP_DIR, f"backup_unobiable_{fecha_str}.xlsx")
        
        try:
            shutil.copy2(self.INPUT_PATH, backup_path)
            print(f"📂 Copia de seguridad creada: {backup_path}")
            return backup_path
        except Exception as e:
            print(f"❌ Error creando backup: {e}")
            return None
    
    def limpiar_procesos_excel(self):
        """Limpia procesos de Excel que puedan estar colgados."""
        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    try:
                        proc.terminate()
                        proc.wait(timeout=5)
                        print(f"🧹 Proceso Excel terminado: PID {proc.info['pid']}")
                    except:
                        pass
        except ImportError:
            print("⚠️ psutil no instalado. Instala con: pip install psutil")
        except Exception as e:
            print(f"⚠️ Error limpiando procesos: {e}")
    
    def remover_solo_lectura(self):
        """Remueve el atributo de solo lectura del archivo."""
        try:
            file_path = Path(self.INPUT_PATH)
            if file_path.exists():
                # Remover atributo de solo lectura
                file_path.chmod(0o666)
                print("🔓 Atributo de solo lectura removido")
        except Exception as e:
            print(f"⚠️ Error removiendo solo lectura: {e}")
    
    def refrescar_conexiones_win32com(self):
        """Actualiza conexiones usando win32com (método alternativo más estable)."""
        print("🔄 Actualizando conexiones con win32com...")
        
        excel = None
        workbook = None
        
        try:
            # Crear instancia de Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            # Abrir el archivo
            workbook = excel.Workbooks.Open(self.INPUT_PATH)
            
            # Actualizar todas las conexiones
            workbook.RefreshAll()
            
            # Esperar a que termine la actualización con timeout
            start_time = time.time()
            while excel.CalculationState != -4208:  # xlDone
                time.sleep(2)
                elapsed = time.time() - start_time
                if elapsed > self.MAX_REFRESH_TIME:
                    print(f"⚠️ Timeout después de {self.MAX_REFRESH_TIME} segundos")
                    break
                print(f"⏳ Actualizando... ({elapsed:.0f}s)")
            
            # Guardar y cerrar
            workbook.Save()
            workbook.Close(SaveChanges=True)
            print("✅ Conexiones actualizadas con win32com")
            return True
            
        except Exception as e:
            print(f"❌ Error con win32com: {e}")
            return False
            
        finally:
            # Limpiar objetos COM
            try:
                if workbook:
                    workbook.Close(SaveChanges=False)
                if excel:
                    excel.Quit()
            except:
                pass
            
            # Liberar referencias COM
            try:
                import gc
                del workbook, excel
                gc.collect()
            except:
                pass
    
    def refrescar_conexiones_xlwings(self):
        """Actualiza conexiones usando xlwings (método original mejorado)."""
        print("🔄 Actualizando conexiones con xlwings...")
        
        app = None
        wb = None
        
        try:
            # Configurar xlwings
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            
            # Abrir archivo
            wb = app.books.open(self.INPUT_PATH)
            
            # Actualizar conexiones
            wb.api.RefreshAll()
            
            # Esperar con timeout mejorado
            start_time = time.time()
            while app.api.CalculationState == 1:  # xlCalculating
                time.sleep(2)
                elapsed = time.time() - start_time
                if elapsed > self.MAX_REFRESH_TIME:
                    print(f"⚠️ Timeout después de {self.MAX_REFRESH_TIME} segundos")
                    break
                print(f"⏳ Calculando... ({elapsed:.0f}s)")
            
            # Guardar
            wb.save()
            wb.close()
            print("✅ Conexiones actualizadas con xlwings")
            return True
            
        except Exception as e:
            print(f"❌ Error con xlwings: {e}")
            return False
            
        finally:
            try:
                if wb:
                    wb.close()
                if app:
                    app.quit()
            except:
                pass
    
    def restaurar_backup(self, backup_path):
        """Restaura el archivo desde el backup si algo sale mal."""
        if backup_path and os.path.exists(backup_path):
            try:
                shutil.copy2(backup_path, self.INPUT_PATH)
                print(f"🔄 Archivo restaurado desde backup: {backup_path}")
                return True
            except Exception as e:
                print(f"❌ Error restaurando backup: {e}")
                return False
        return False
    
    def main(self):
        print("🚀 Iniciando actualización de UnoBiable...")
        
        # Verificaciones iniciales
        if not os.path.exists(self.INPUT_PATH):
            print(f"❌ Archivo no encontrado: {self.INPUT_PATH}")
            return
        
        # 1️⃣ Limpiar procesos Excel previos
        print("🧹 Limpiando procesos Excel...")
        self.limpiar_procesos_excel()
        time.sleep(2)
        
        # 2️⃣ Verificar que el archivo esté disponible
        if not self.verificar_archivo_disponible():
            print("❌ El archivo está siendo usado por otro proceso")
            return
        
        # 3️⃣ Remover solo lectura
        self.remover_solo_lectura()
        
        # 4️⃣ Crear backup
        backup_path = self.hacer_backup()
        if not backup_path:
            print("❌ No se pudo crear backup. Abortando.")
            return
        
        # 5️⃣ Intentar actualizar con diferentes métodos
        success = False
        
        for attempt in range(self.RETRY_ATTEMPTS):
            print(f"\n📋 Intento {attempt + 1} de {self.RETRY_ATTEMPTS}")
            
            # Primero intentar con win32com (más estable)
            if self.refrescar_conexiones_win32com():
                success = True
                break
            
            print("⚠️ win32com falló, intentando con xlwings...")
            time.sleep(3)
            
            # Si falla, intentar con xlwings
            if self.refrescar_conexiones_xlwings():
                success = True
                break
                
            # Esperar antes del siguiente intento
            if attempt < self.RETRY_ATTEMPTS - 1:
                print(f"⏳ Esperando antes del siguiente intento...")
                time.sleep(5)
        
        # 6️⃣ Verificar resultado
        if success:
            print("🎉 ¡Actualización completada exitosamente!")
            
            # Verificar que el archivo no quedó en solo lectura
            if self.verificar_archivo_disponible():
                print("✅ El archivo está disponible para edición")
            else:
                print("⚠️ El archivo podría estar en solo lectura")
                self.remover_solo_lectura()
                
        else:
            print("❌ Falló la actualización después de todos los intentos")
            print("🔄 Restaurando desde backup...")
            
            if self.restaurar_backup(backup_path):
                print("✅ Archivo restaurado exitosamente")
            else:
                print("❌ Error restaurando archivo")
        
        # 7️⃣ Limpieza final
        print("🧹 Limpieza final...")
        self.limpiar_procesos_excel()
        print("🏁 Proceso terminado")

if __name__ == "__main__":
    updater = UnoBiableUpdater()
    updater.main()
    