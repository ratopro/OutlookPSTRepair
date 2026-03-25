#!/usr/bin/env python3
"""
Versión simple de la herramienta de reparación PST que funciona sin interacción
Muestra el árbol de carpetas y permite especificar selección vía argumentos
"""

import pypff
import os
import sys

# Códigos de color ANSI
class Colors:
    GREEN = '\033[92m'      # Correcto
    YELLOW = '\033[93m'     # Dañado
    RED = '\033[91m'        # Eliminado
    BLUE = '\033[94m'       # Selección
    END = '\033[0m'         # Fin de color

def supports_color():
    """Verifica si la terminal soporta colores"""
    return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

def colorize(text, color):
    """Aplica color al texto si la terminal lo soporta"""
    if not supports_color():
        return text
    return f"{color}{text}{Colors.END}"

class PSTRepairSimple:
    def __init__(self):
        self.pst_file = None
        self.folders_data = []
        self.selected_for_recovery = set()
        self._last_pst_path = None
    
    def analyze_pst(self, pst_path):
        """Analiza el archivo PST"""
        if not os.path.exists(pst_path):
            print(f"Error: El archivo {pst_path} no existe")
            return False
        
        try:
            print(f"Abriendo archivo PST: {os.path.basename(pst_path)}")
            self.pst_file = pypff.file()
            self.pst_file.open(pst_path)
            
            # Información básica
            file_size = os.path.getsize(pst_path)
            print(f"Tamaño: {file_size / (1024*1024):.2f} MB")
            print("Analizando estructura de carpetas...")
            
            # Limpiar y analizar
            self.folders_data = []
            self.selected_for_recovery.clear()
            
            root_folder = self.pst_file.get_root_folder()
            self.process_folder(root_folder, "", 0, is_root=True)
            
            self.pst_file.close()
            return True
            
        except Exception as e:
            print(f"Error al analizar el PST: {e}")
            if self.pst_file:
                try:
                    self.pst_file.close()
                except:
                    pass
            return False
    
    def process_folder(self, folder, parent_path, depth, is_root=False):
        """Procesa una carpeta y sus subcarpetas"""
        try:
            # Nombre de la carpeta
            if is_root:
                folder_name = "<Carpeta Raíz>"
                display_indent = ""
                full_path = "<Carpeta Raíz>"
            else:
                folder_name = folder.get_name()
                if not folder_name:
                    folder_name = "<Sin nombre>"
                display_indent = "    " * depth
                if parent_path:
                    full_path = f"{parent_path}\\{folder_name}"
                else:
                    full_path = folder_name
            
            # Conteos
            try:
                message_count = folder.number_of_sub_messages
            except:
                message_count = 0
            
            # Estado
            status = self.get_folder_status(folder)
            
            # Almacenar
            folder_info = {
                'folder_obj': folder if not is_root else None,
                'name': folder_name,
                'full_path': full_path,
                'indent': display_indent,
                'depth': depth,
                'emails': message_count,
                'status': status,
                'parent': parent_path if not is_root else None,
                'is_root': is_root
            }
            
            self.folders_data.append(folder_info)
            
            # Procesar subcarpetas
            try:
                sub_count = folder.number_of_sub_folders
                for i in range(sub_count):
                    sub_folder = folder.get_sub_folder(i)
                    child_parent = full_path if not is_root else folder_name
                    child_depth = depth + 1 if not is_root else 1
                    self.process_folder(sub_folder, child_parent, child_depth)
            except:
                pass
            else:
                # Para raíz, procesamos sus hijas
                try:
                    sub_count = folder.number_of_sub_folders
                    for i in range(sub_count):
                        sub_folder = folder.get_sub_folder(i)
                        self.process_folder(sub_folder, folder_name, 1)
                except:
                    pass
                    
        except Exception as e:
            # Carpeta dañada
            if is_root:
                folder_name = "<Carpeta Raíz>"
                display_indent = ""
                full_path = "<Carpeta Raíz>"
            else:
                folder_name = "<Error>"
                display_indent = "    " * depth
                if parent_path:
                    full_path = f"{parent_path}\\{folder_name}"
                else:
                    full_path = folder_name
            
            folder_info = {
                'folder_obj': None,
                'name': folder_name,
                'full_path': full_path,
                'indent': display_indent,
                'depth': depth,
                'emails': 0,
                'status': 'damaged',
                'parent': parent_path if not is_root else None,
                'is_root': is_root
            }
            
            self.folders_data.append(folder_info)
    
    def get_folder_status(self, folder):
        """Determina el estado de la carpeta"""
        try:
            name = folder.get_name()
            
            # Heurística para eliminadas
            if name and any(word in name.lower() for word in ['deleted', 'eliminado', 'paper', 'basur', 'trash', 'deleted items']):
                return 'deleted'
            
            # Verificar integridad
            _ = folder.number_of_sub_folders
            _ = folder.number_of_sub_messages
            
            if folder.number_of_sub_messages > 0:
                try:
                    _ = folder.get_sub_message(0)
                except:
                    if folder.number_of_sub_messages > 0:
                        return 'damaged'
            
            return 'correct'
            
        except:
            return 'damaged'
    
    def display_tree(self):
        """Muestra el árbol de carpetas con colores"""
        print("\n" + "="*70)
        print("ÁRBOL DE CARPETAS PST")
        print("="*70)
        
        # Leyenda
        print("Estados: " + 
              colorize("✓ Correcto", Colors.GREEN) + " | " +
              colorize("⚠ Dañado", Colors.YELLOW) + " | " +
              colorize("✗ Eliminado", Colors.RED))
        print("-"*70)
        
        # Mostrar raíz primero
        roots = [f for f in self.folders_data if f.get('is_root', False)]
        for folder in roots:
            self._print_folder(folder, is_last=True, prefix="")
        
        # Luego resto ordenado por ruta
        others = [f for f in self.folders_data if not f.get('is_root', False)]
        others.sort(key=lambda x: x['full_path'])
        
        for folder in others:
            self._print_folder(folder, is_last=False, prefix="")
        
        print("-"*70)
        total_folders = len(self.folders_data)
        total_emails = sum(f['emails'] for f in self.folders_data)
        
        print(f"Total: {total_folders} carpetas, {total_emails} correos")
        print("="*70)
    
    def _print_folder(self, folder, is_last=False, prefix=""):
        """Imprime un ítem de carpeta en el árbol"""
        # Símbolos de árbol
        if prefix == "":
            connector = ""
        else:
            connector = "└── " if is_last else "├── "
        
        # Símbolo de estado con color
        status_symbols = {
            'correct': ('✓', Colors.GREEN),
            'damaged': ('⚠', Colors.YELLOW),
            'deleted': ('✗', Colors.RED)
        }
        symbol, color = status_symbols.get(folder['status'], ('?', Colors.END))
        status_display = colorize(symbol, color)
        
        # Texto de la carpeta
        display_text = f"{prefix}{connector}{status_display} {folder['indent']}{folder['name']} ({folder['emails']} correos)"
        print(display_text)
        
        # Prefix para hijos
        if prefix == "":
            new_prefix = ""
        else:
            extension = "    " if is_last else "│   "
            new_prefix = prefix + extension
        
        # Hijos
        children = [f for f in self.folders_data 
                   if f.get('parent') == folder['full_path'] 
                   and not f.get('is_root', False)]
        children.sort(key=lambda x: x['name'])
        
        for i, child in enumerate(children):
            is_last_child = (i == len(children) - 1)
            self._print_folder(child, is_last_child, new_prefix)
    
    def select_folders(self, selection_str):
        """Selecciona carpetas basado en una cadena de selección"""
        if not selection_str or selection_str.lower() == 'none':
            self.selected_for_recovery.clear()
            return
        
        if selection_str.lower() == 'all':
            self.selected_for_recovery = {f['full_path'] for f in self.folders_data}
            return
        
        # Procesar selección
        self.selected_for_recovery.clear()
        parts = [p.strip() for p in selection_str.split(',') if p.strip()]
        
        for part in parts:
            if '-' in part:
                # Rango
                try:
                    start, end = map(int, part.split('-'))
                    start = max(1, start)
                    end = min(len(self.folders_data), end)
                    for i in range(start, end + 1):
                        self.selected_for_recovery.add(self.folders_data[i-1]['full_path'])
                except ValueError:
                    print(f"Advertencia: Rango inválido '{part}'")
            else:
                # Número simple
                try:
                    idx = int(part)
                    if 1 <= idx <= len(self.folders_data):
                        self.selected_for_recovery.add(self.folders_data[idx-1]['full_path'])
                    else:
                        print(f"Advertencia: Índice fuera de rango '{part}'")
                except ValueError:
                    print(f"Advertencia: Entrada inválida '{part}'")
    
    def display_selected_tree(self):
        """Muestra el árbol indicando qué está seleccionado"""
        print("\n" + "="*70)
        print("ÁRBOL DE CARPETAS PST - SELECCIÓN PARA RECUPERACIÓN")
        print("="*70)
        
        # Leyenda
        print("Estados: " + 
              colorize("✓ Correcto", Colors.GREEN) + " | " +
              colorize("⚠ Dañado", Colors.YELLOW) + " | " +
              colorize("✗ Eliminado", Colors.RED))
        print("Selección: " + 
              colorize("[x] Seleccionado", Colors.BLUE) + " | " +
              colorize("[ ] No seleccionado", Colors.BLUE))
        print("-"*70)
        
        # Mostrar raíz primero
        roots = [f for f in self.folders_data if f.get('is_root', False)]
        for folder in roots:
            self._print_selected_folder(folder, is_last=True, prefix="")
        
        # Luego resto ordenado por ruta
        others = [f for f in self.folders_data if not f.get('is_root', False)]
        others.sort(key=lambda x: x['full_path'])
        
        for folder in others:
            self._print_selected_folder(folder, is_last=False, prefix="")
        
        print("-"*70)
        total_folders = len(self.folders_data)
        total_emails = sum(f['emails'] for f in self.folders_data)
        selected_folders = len(self.selected_for_recovery)
        selected_emails = sum(f['emails'] for f in self.folders_data 
                           if f['full_path'] in self.selected_for_recovery)
        
        print(f"Total: {total_folders} carpetas, {total_emails} correos")
        print(f"Seleccionadas: {selected_folders} carpetas, {selected_emails} correos")
        print("="*70)
    
    def _print_selected_folder(self, folder, is_last=False, prefix=""):
        """Imprime un ítem de carpeta mostrando su estado de selección"""
        # Símbolos de árbol
        if prefix == "":
            connector = ""
        else:
            connector = "└── " if is_last else "├── "
        
        # Símbolo de estado con color
        status_symbols = {
            'correct': ('✓', Colors.GREEN),
            'damaged': ('⚠', Colors.YELLOW),
            'deleted': ('✗', Colors.RED)
        }
        symbol, color = status_symbols.get(folder['status'], ('?', Colors.END))
        status_display = colorize(symbol, color)
        
        # Checkbox de selección
        is_selected = folder['full_path'] in self.selected_for_recovery
        checkbox = colorize("[x]", Colors.BLUE) if is_selected else colorize("[ ]", Colors.BLUE)
        
        # Texto de la carpeta
        display_text = f"{prefix}{connector}{checkbox} {status_display} {folder['indent']}{folder['name']} ({folder['emails']} correos)"
        print(display_text)
        
        # Prefix para hijos
        if prefix == "":
            new_prefix = ""
        else:
            extension = "    " if is_last else "│   "
            new_prefix = prefix + extension
        
        # Hijos
        children = [f for f in self.folders_data 
                   if f.get('parent') == folder['full_path'] 
                   and not f.get('is_root', False)]
        children.sort(key=lambda x: x['name'])
        
        for i, child in enumerate(children):
            is_last_child = (i == len(children) - 1)
            self._print_selected_folder(child, is_last_child, new_prefix)
    
    def run_recovery_simulation(self, output_path):
        """Simula el proceso de recuperación"""
        if not self.selected_for_recovery:
            print("\nNo hay carpetas seleccionadas para recuperar")
            return False
        
        print(f"\nSimulando recuperación a: {output_path}")
        print("Selección actual:")
        
        selected_folders = [f for f in self.folders_data if f['full_path'] in self.selected_for_recovery]
        selected_folders.sort(key=lambda x: x['full_path'])
        
        total_emails = 0
        for folder in selected_folders:
            status_sym = {'correct': '✓', 'damaged': '⚠', 'deleted': '✗'}[folder['status']]
            indent = "    " * folder['depth']
            marker = "" if folder.get('is_root', False) else indent
            print(f"  {status_sym} {marker}{folder['name']} ({folder['emails']} correos)")
            total_emails += folder['emails']
        
        print(f"\nTotal de correos a recuperar: {total_emails}")
        print("\nNota: Esta es una simulación. En una implementación real:")
        print("1. Se crearía un nuevo archivo PST")
        print("2. Se copiarían solo las carpetas seleccionadas con sus contenidos")
        print("3. Se preservaría la estructura jerárquica")
        print("4. Se mantendrían los metadatos de los mensajes")
        
        # Simular creación del archivo
        try:
            # En realidad, aquí iría la lógica de extracción
            # Por ahora, solo copiamos el archivo como ejemplo
            import shutil
            if os.path.exists(self._last_pst_path):
                shutil.copy2(self._last_pst_path, output_path)
                print(f"\nArchivo de simulación creado: {output_path}")
                print(f"Tamaño: {os.path.getsize(output_path) / (1024*1024):.2f} MB")
                return True
            else:
                print("Error: Archivo origen no disponible para copia")
                return False
        except Exception as e:
            print(f"Error al crear archivo de simulación: {e}")
            return False
    
    def run(self, pst_path, selection_str=None, output_path=None):
        """Ejecuta el proceso completo"""
        print("Herramienta de Análisis de PST de Outlook")
        print("="*45)
        
        # Guardar ruta para posible recuperación
        self._last_pst_path = pst_path
        
        # Analizar
        if not self.analyze_pst(pst_path):
            return False
        
        # Mostrar árbol inicial
        self.display_tree()
        
        # Procesar selección si se proporcionó
        if selection_str is not None:
            print(f"\nAplicando selección: {selection_str}")
            self.select_folders(selection_str)
            self.display_selected_tree()
            
            # Si se especificó output, intentar recuperación
            if output_path:
                self.run_recovery_simulation(output_path)
        
        return True

def main():
    if len(sys.argv) < 2 or (len(sys.argv) > 1 and sys.argv[1] in ('--help', '-h', '/?')):
        print("=" * 60)
        print("  HERRAMIENTA DE REPARACIÓN DE ARCHIVOS PST DE OUTLOOK")
        print("=" * 60)
        print("")
        print("Uso: OutlookPSTRepair.exe <archivo.pst> [selección] [salida]")
        print("")
        print("Argumentos:")
        print("  archivo.pst   Ruta al archivo PST a analizar (obligatorio)")
        print("  selección     (Opcional) Carpetas a seleccionar:")
        print("                  all    - Todas las carpetas")
        print("                  none   - Ninguna carpeta")
        print("                  1,3,5  - Carpetas específicas por número")
        print("                  1-5    - Rango de carpetas")
        print("                  1,3-5,7 - Combinación")
        print("  salida        (Opcional) Ruta para guardar el PST recuperado")
        print("")
        print("Ejemplos:")
        print("  OutlookPSTRepair.exe C:\\ruta\\archivo.pst")
        print("  OutlookPSTRepair.exe archivo.pst all")
        print("  OutlookPSTRepair.exe archivo.pst 1,3,5 recuperado.pst")
        print("  OutlookPSTRepair.exe archivo.pst 2-4")
        print("=" * 60)
        sys.exit(1)
    
    pst_path = sys.argv[1]
    selection_str = sys.argv[2] if len(sys.argv) > 2 else None
    output_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    # Expandir rutas
    pst_path = os.path.expanduser(pst_path)
    pst_path = os.path.expandvars(pst_path)
    
    if output_path:
        output_path = os.path.expanduser(output_path)
        output_path = os.path.expandvars(output_path)
    
    tool = PSTRepairSimple()
    success = tool.run(pst_path, selection_str, output_path)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()