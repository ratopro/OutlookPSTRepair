#!/usr/bin/env python3
"""
Aplicación para reparar archivos PST de Outlook con vista dividida y paneles redimensionables
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pypff
import os

class PSTRepairApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Reparador de Archivos PST de Outlook")
        self.root.geometry("1100x700")
        
        # Variables
        self.pst_file_path = tk.StringVar()
        self.pst_file = None
        self.folders_data = []
        self.selected_items = set()
        self._folder_item_by_tree_id = {}
        self._tree_node_by_path = {}
        
        self.setup_ui()
        
        # Cerrar archivo PST al cerrar la aplicación
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def on_closing(self):
        """Cierra el archivo PST y termina la aplicación"""
        if self.pst_file:
            try:
                self.pst_file.close()
            except:
                pass
        self.root.destroy()
    
    def setup_ui(self):
        """Configura la interfaz de usuario"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Selección de archivo PST
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="Archivo PST:").pack(side=tk.LEFT)
        ttk.Entry(file_frame, textvariable=self.pst_file_path, width=70).pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Examinar...", command=self.browse_pst).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(file_frame, text="Analizar", command=self.analyze_pst).pack(side=tk.LEFT, padx=(5, 0))
        
        # Panel de información
        info_frame = ttk.LabelFrame(main_frame, text="Información del Archivo", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_size_label = ttk.Label(info_frame, text="Tamaño: N/A")
        self.file_size_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        self.folders_count_label = ttk.Label(info_frame, text="Carpetas: N/A")
        self.folders_count_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        self.messages_count_label = ttk.Label(info_frame, text="Mensajes: N/A")
        self.messages_count_label.grid(row=0, column=2, sticky=tk.W, padx=(0, 20))
        
        self.progress = ttk.Progressbar(info_frame, mode='indeterminate', length=200)
        self.progress.grid(row=0, column=3, sticky=tk.E, padx=(20, 0))
        self.progress.grid_remove()
        
        # PanedWindow para paneles redimensionables
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Frame izquierdo: Árbol de carpetas
        left_frame = ttk.LabelFrame(paned, text="Estructura de Carpetas", padding="5")
        paned.add(left_frame, weight=3)
        
        # Toolbar con botones de expandir/colapsar
        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(toolbar, text="Desplegar", command=self.expand_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Colapsar", command=self.collapse_selected).pack(side=tk.LEFT, padx=5)
        
        # Crear Treeview con scrollbar - estilo de árbol expandible
        tree_container = ttk.Frame(left_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Treeview con estilo de árbol
        self.tree = ttk.Treeview(tree_container, show="tree headings")
        
        # Columnas
        self.tree["columns"] = ("select", "emails", "status")
        
        # Encabezados
        self.tree.heading("#0", text="Carpeta", anchor=tk.W)
        self.tree.heading("select", text="✓", anchor=tk.CENTER)
        self.tree.heading("emails", text="Correos", anchor=tk.CENTER)
        self.tree.heading("status", text="Estado", anchor=tk.CENTER)
        
        # Ancho de columnas
        self.tree.column("#0", width=400)
        self.tree.column("select", width=40, anchor=tk.CENTER)
        self.tree.column("emails", width=80, anchor=tk.CENTER)
        self.tree.column("status", width=100, anchor=tk.CENTER)
        
        # Scrollbar
        tree_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Habilitar selección de carpeta
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        # Habilitar clic en columna de selección para toggle de carpeta
        self.tree.bind("<Button-1>", self.on_tree_click)
        # Legend de estado de carpetas (faltantes, etc.)
        legend_frame = ttk.Frame(left_frame)
        legend_frame.pack(fill=tk.X, pady=(4, 8))
        def _legend_item(text, color):
            sw = tk.Label(legend_frame, width=2, bg=color)
            sw.pack(side=tk.LEFT, padx=(4, 2))
            lbl = tk.Label(legend_frame, text=text)
            lbl.pack(side=tk.LEFT, padx=(0, 6))
        _legend_item("Correcto", "#28a745")
        _legend_item("Dañado", "#ff9800")
        _legend_item("Eliminado", "#f44336")
        _legend_item("Missing", "#ffd6f7")
        
        # Configurar colores para estados
        self.tree.tag_configure("correct", foreground="green")
        self.tree.tag_configure("damaged", foreground="orange")
        self.tree.tag_configure("deleted", foreground="red")
        self.tree.tag_configure("selected", background="#e3f2fd")
        self.tree.tag_configure("missing", foreground="black", background="#ffd6f7")
        
        # Botones de acción
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(button_frame, text="Seleccionar Todo", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Deseleccionar Todo", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reparar Seleccionadas", command=self.repair_selected).pack(side=tk.LEFT, padx=20)
        ttk.Button(button_frame, text="Salir", command=self.root.quit).pack(side=tk.RIGHT, padx=5)
        
        # Barra de estado
        self.status_var = tk.StringVar()
        self.status_var.set("Listo para cargar archivo PST")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X)
    
    def browse_pst(self):
        """Abre el diálogo para seleccionar un archivo PST"""
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo PST",
            filetypes=[("Archivos PST", "*.pst"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.pst_file_path.set(filename)
            self.status_var.set(f"Archivo seleccionado: {os.path.basename(filename)}")
    
    def analyze_pst(self):
        """Analiza el archivo PST seleccionado"""
        pst_path = os.path.normpath(self.pst_file_path.get())
        if not pst_path:
            messagebox.showerror("Error", "Por favor seleccione un archivo PST primero")
            return
        
        if not os.path.exists(pst_path):
            messagebox.showerror("Error", "El archivo PST no existe")
            return
        
        self.progress.grid()
        self.progress.config(mode='determinate', maximum=100, value=0)
        self.progress.start()
        
        try:
            self.status_var.set("Abriendo archivo PST...")
            self.root.update_idletasks()
            
            self.pst_file = pypff.file()
            self.pst_file.open(pst_path)
            self.progress.step(20)
            self.root.update_idletasks()
            
            file_size = os.path.getsize(pst_path)
            self.file_size_label.config(text=f"Tamaño: {file_size / (1024*1024):.2f} MB")
            
            self.status_var.set("Analizando estructura de carpetas...")
            self.root.update_idletasks()
            
            self.folders_data = []
            self.selected_items.clear()
            
            root_folder = self.pst_file.get_root_folder()
            self.analyze_folder_recursive(root_folder, "", 0)
            self.progress.step(60)
            self.root.update_idletasks()
            
            total_folders = len(self.folders_data)
            total_messages = sum(folder['message_count'] for folder in self.folders_data)
            self.folders_count_label.config(text=f"Carpetas: {total_folders}")
            self.messages_count_label.config(text=f"Mensajes: {total_messages}")
            
            self.populate_tree()
            
            self.progress.stop()
            self.progress.config(mode='indeterminate')
            self.progress.grid_remove()
            self.status_var.set(f"Análisis completado. Se encontraron {total_folders} carpetas.")
            messagebox.showinfo("Éxito", f"Análisis completado.\nCarpetas encontradas: {total_folders}\nMensajes totales: {total_messages}")
            
        except Exception as e:
            self.progress.stop()
            self.progress.grid_remove()
            messagebox.showerror("Error", f"No se pudo analizar el archivo PST:\n{str(e)}")
            self.status_var.set("Error al analizar el archivo")
            import traceback
            traceback.print_exc()
    
    def analyze_folder_recursive(self, folder, parent_path, depth):
        """Analiza recursivamente una carpeta y sus subcarpetas"""
        try:
            # Obtener nombre de la carpeta
            folder_name = folder.get_name()
            if not folder_name:
                folder_name = "<Sin nombre>"
            
            # Construir ruta completa
            if parent_path:
                full_path = f"{parent_path}\\{folder_name}"
            else:
                full_path = folder_name
            
            # Obtener número de mensajes
            try:
                message_count = folder.number_of_sub_messages
            except:
                message_count = 0
            
            # Determinar estado de la carpeta
            status = self.determine_folder_status(folder)
            
            # Almacenar información
            folder_info = {
                'folder_obj': folder,
                'folder_name': folder_name,
                'full_path': full_path,
                'depth': depth,
                'message_count': message_count,
                'status': status,
                'parent_path': parent_path if parent_path else None
            }
            
            self.folders_data.append(folder_info)
            
            # Procesar subcarpetas
            try:
                sub_folder_count = folder.number_of_sub_folders
                for i in range(sub_folder_count):
                    sub_folder = folder.get_sub_folder(i)
                    self.analyze_folder_recursive(sub_folder, full_path, depth + 1)
            except:
                pass
                
        except Exception as e:
            # En caso de error al leer la carpeta
            folder_name = "<Error>"
            
            if parent_path:
                full_path = f"{parent_path}\\{folder_name}"
            else:
                full_path = folder_name
            
            folder_info = {
                'folder_obj': None,
                'folder_name': folder_name,
                'full_path': full_path,
                'depth': depth,
                'message_count': 0,
                'status': 'damaged',
                'parent_path': parent_path if parent_path else None
            }
            
            self.folders_data.append(folder_info)
    
    def determine_folder_status(self, folder):
        """Determina el estado de una carpeta"""
        try:
            name = folder.get_name()
            
            # Verificar si es una carpeta eliminada
            if name and any(word in name.lower() for word in ['deleted', 'eliminado', 'paper', 'basur', 'trash', 'deleted items']):
                return "deleted"
            
            # Verificar acceso a propiedades
            _ = folder.number_of_sub_folders
            _ = folder.number_of_sub_messages
            
            # Verificar integridad de mensajes
            if folder.number_of_sub_messages > 0:
                try:
                    _ = folder.get_sub_message(0)
                except:
                    if folder.number_of_sub_messages > 0:
                        return "damaged"
            
            return "correct"
            
        except:
            return "damaged"
    
    def populate_tree(self):
        """Pobla el árbol de carpetas en la interfaz"""
        # Limpiar el árbol
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Limpiar referencias de tree_item
        for f in self.folders_data:
            f['tree_item'] = ""
        
        # Crear diccionario de carpetas por path para facilitar búsqueda
        folders_dict = {f['full_path']: f for f in self.folders_data}
        
        # Insertar carpetas en orden de profundidad
        self._folder_item_by_tree_id = {}
        for depth in range(max((f['depth'] for f in self.folders_data), default=0) + 1):
            for folder_info in sorted(self.folders_data, key=lambda x: x['depth']):
                if folder_info['depth'] == depth:
                    self.insert_folder_item(folder_info, folders_dict)
    
    def insert_folder_item(self, folder_info, folders_dict):
        """Inserta un ítem de carpeta en el árbol"""
        # Preparar valores
        select_val = "☑" if folder_info['full_path'] in self.selected_items else "☐"
        emails_val = str(folder_info['message_count'])
        status_val = folder_info['status'].capitalize()
        
        # Determinar tag para color
        tags = [folder_info['status']]
        if folder_info['full_path'] in self.selected_items:
            tags.append("selected")
        
        # Determinar el parent_id correcto
        parent_id = ""
        # Asegurar que el padre exista en la jerarquía; si no, crear un nodo missing (solo para visual)
        if folder_info['parent_path']:
            if folder_info['parent_path'] in folders_dict:
                parent_folder = folders_dict[folder_info['parent_path']]
                parent_id = parent_folder.get('tree_item', '')
            else:
                # Crear padre missing si no existe
                parent_path = folder_info['parent_path']
                if not any(f['full_path'] == parent_path for f in self.folders_data):
                    parent_name = parent_path.split('\\')[-1]
                    parent_depth = max(0, folder_info['depth'] - 1)
                    parent_missing = {
                        'folder_obj': None,
                        'folder_name': parent_name,
                        'full_path': parent_path,
                        'depth': parent_depth,
                        'message_count': 0,
                        'status': 'missing',
                        'parent_path': None,
                        'is_root': False,
                        'tree_item': None
                    }
                    self.folders_data.append(parent_missing)
                    folders_dict[parent_path] = parent_missing
                    # Crear nodo visual para el padre missing si no existe todavía en el árbol
                    try:
                        missing_text = parent_name
                        missing_id = self.tree.insert("", tk.END, text=missing_text, values=("☐", 0, "Missing"), tags=("missing",))
                        parent_missing['tree_item'] = missing_id
                        self._folder_item_by_tree_id[missing_id] = parent_missing
                    except Exception:
                        pass
                if folders_dict.get(parent_path) and folders_dict[parent_path].get('tree_item'):
                    parent_id = folders_dict[parent_path]['tree_item']
                else:
                    parent_id = ""
        
        # Si el padre no existe en la lista, crear un padre faltante (missing)
        # Insertar ítem (sin crear nodos missing)
        try:
            # Determinar parent_id real basado en parent_path si existe entre folders_data
            if folder_info['parent_path']:
                parent = next((f for f in self.folders_data if f['full_path'] == folder_info['parent_path']), None)
                parent_id = parent['tree_item'] if parent and 'tree_item' in parent else ""
            else:
                parent_id = ""
            item_id = self.tree.insert(
                parent_id, 
                tk.END, 
                text=folder_info['folder_name'],
                values=(select_val, emails_val, status_val),
                tags=tuple(tags),
                open=False  # Iniciar colapsado
            )
            # Guardar referencia
            folder_info['tree_item'] = item_id
            self._folder_item_by_tree_id[item_id] = folder_info
        except Exception as e:
            print(f"Error inserting folder {folder_info['folder_name']}: {e}")
    
    def on_tree_select(self, event):
        """Muestra los correos de la carpeta seleccionada"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        
        # Intentar obtener la carpeta desde el mapeo por ID de árbol
        folder_info = self._folder_item_by_tree_id.get(item)
        if folder_info is None:
            folder_info = None
            for f in self.folders_data:
                if f.get('tree_item') == item:
                    folder_info = f
                    break
        
        if not folder_info:
            # Carpeta no encontrada en la lista (posible carpeta faltante)
            self.email_tree.delete(*self.email_tree.get_children())
            self.email_info_label.config(text="Carpeta no encontrada (faltante para recuperación)")
            return
        if not folder_info['folder_obj']:
            if folder_info.get('status') == 'missing':
                self.email_tree.delete(*self.email_tree.get_children())
                self.email_info_label.config(text=f"Carpeta faltante: {folder_info['folder_name']}")
            else:
                self.email_info_label.config(text="Carpeta no accesible")
            return
        
        # Limpiar treeview de correos
        for item in self.email_tree.get_children():
            self.email_tree.delete(item)
        
        # Limpiar selección de correos
        self.selected_emails.clear()
        
        try:
            folder_obj = folder_info['folder_obj']
            message_count = folder_obj.number_of_sub_messages

            self.email_info_label.config(text=f"Carpeta: {folder_info['folder_name']} - {message_count} correos")

            if message_count == 0:
                # Limpiar lista de correos si no hay
                self.email_tree.delete(*self.email_tree.get_children())
                self._email_item_map.clear()
                self._current_email_folder = folder_info['full_path']
                return

            # Preparar lista de correos en email_tree
            self.email_tree.delete(*self.email_tree.get_children())
            self._email_item_map.clear()
            self._current_email_folder = folder_info['full_path']
            for i in range(message_count):
                try:
                    msg = folder_obj.get_sub_message(i)
                    
                    # Obtener propiedades del mensaje
                    subject = None
                    if hasattr(msg, "get_property_value"):
                        try:
                            subject = msg.get_property_value(0x0037)
                        except Exception:
                            subject = None
                    if not subject:
                        subject = "(Sin asunto)"
                    
                    sender = None
                    if hasattr(msg, "get_property_value"):
                        for key in (0x0C1A, 0x0C1B, 0x0C1D, 0x0C1F, 0x0C22):
                            try:
                                val = msg.get_property_value(key)
                                if val:
                                    sender = val
                                    break
                            except Exception:
                                pass
                    if not sender:
                        sender = "(Desconocido)"

                    # Fecha, con múltiples fallbacks
                    date = None
                    if hasattr(msg, "get_property_value"):
                        for key in (0x0E06, 0x0E07, 0x0E08, 0x3A0):
                            try:
                                val = msg.get_property_value(key)
                                if val:
                                    date = val
                                    break
                            except Exception:
                                pass
                    date_str = str(date)[:10] if date else ""
                    
                    # Insertar en el treeview de correos
                    item_id = self.email_tree.insert(
                        "",
                        tk.END,
                        values=("☐", sender, subject, date_str),
                        tags=("unselected",)
                    )
                    self._email_item_map[item_id] = (folder_info['full_path'], i)
                except Exception:
                    item_id = self.email_tree.insert(
                        "",
                        tk.END,
                        values=("☐", "(Desconocido)", "(No leer)", ""),
                        tags=("error",)
                    )
                    self._email_item_map[item_id] = (folder_info['full_path'], i)
        except Exception as e:
            self.email_info_label.config(text=f"Error: {str(e)}")
    
    def on_email_click(self, event):
        """Maneja clic en checkbox de correo"""
        region = self.email_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.email_tree.identify_column(event.x)
        if column != "#1":  # Solo la primera columna (checkbox)
            return
        
        item = self.email_tree.identify_row(event.y)
        if not item:
            return
        # Obtener valores actuales
        values = list(self.email_tree.item(item, 'values'))
        current_check = values[0]
        # Mapping para el correo actual
        key = self._email_item_map.get(item)
        if key:
            folder, idx = key
            if (folder, idx) in self.selected_emails:
                self.selected_emails.remove((folder, idx))
                values[0] = "☐"
            else:
                self.selected_emails.add((folder, idx))
                values[0] = "☑"
            self.email_tree.item(item, values=tuple(values))
        
        # Actualizar contador
        self.update_email_count()
    
    def update_email_count(self):
        """Actualiza el conteo de correos seleccionados"""
        # Conteo dentro de la carpeta actual
        current = self._current_email_folder
        total_count = 0
        selected_count = 0
        for item, (fld, idx) in list(self._email_item_map.items()):
            if fld == current:
                total_count += 1
                if (fld, idx) in self.selected_emails:
                    selected_count += 1
        folder_name = ''
        if current:
            folder_info = next((f for f in self.folders_data if f['full_path'] == current), None)
            if folder_info:
                folder_name = folder_info['folder_name']
        self.email_info_label.config(text=f"{folder_name} - {selected_count}/{total_count} seleccionados")

    def on_tree_click(self, event):
        """Maneja clic en la columna de selección de carpetas para hacer toggle"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col not in ("#1", "1"):  # primera columna de datos (selección)
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        # Obtener carpeta asociada
        folder_info = self._folder_item_by_tree_id.get(item)
        if folder_info is None:
            folder_info = next((f for f in self.folders_data if f.get('tree_item') == item), None)
        if not folder_info:
            return
        # Toggle
        if folder_info['full_path'] in self.selected_items:
            self.selected_items.remove(folder_info['full_path'])
            self.tree.set(item, "select", "☐")
        else:
            self.selected_items.add(folder_info['full_path'])
            self.tree.set(item, "select", "☑")
        self.status_var.set(f"{len(self.selected_items)} de {len(self.folders_data)} carpetas seleccionadas")
        # Cascade selection to all descendants
        is_selected = folder_info['full_path'] in self.selected_items
        self._cascade_select(folder_info['full_path'], is_selected)

    def _collect_descendant_paths(self, folder_path):
        """Recorre recursivamente para obtener todos los descendientes (sin incluir la carpeta raiz)"""
        descendants = []
        stack = [folder_path]
        while stack:
            current = stack.pop()
            for f in self.folders_data:
                if f.get('parent_path') == current:
                    descendants.append(f['full_path'])
                    stack.append(f['full_path'])
        return descendants

    def _cascade_select(self, folder_path, select_bool):
        """Propaga la selección/deselección a las subcarpetas descendientes"""
        for path in self._collect_descendant_paths(folder_path):
            if select_bool:
                self.selected_items.add(path)
            else:
                self.selected_items.discard(path)
            f = next((ff for ff in self.folders_data if ff['full_path'] == path), None)
            if f and f.get('tree_item'):
                self.tree.set(f['tree_item'], "select", "☑" if select_bool else "☐")
    
    def toggle_all_selection(self):
        """Alterna selección de todos los ítems"""
        if len(self.selected_items) == len(self.folders_data):
            self.selected_items.clear()
        else:
            self.selected_items = {f['full_path'] for f in self.folders_data}
        
        self.refresh_tree_checkboxes()
        self.status_var.set(f"{len(self.selected_items)} de {len(self.folders_data)} carpetas seleccionadas")
    
    def select_all(self):
        """Selecciona todas las carpetas"""
        self.selected_items = {f['full_path'] for f in self.folders_data}
        self.refresh_tree_checkboxes()
        self.status_var.set(f"{len(self.selected_items)} carpetas seleccionadas")
    
    def deselect_all(self):
        """Deselecciona todas las carpetas"""
        self.selected_items.clear()
        self.refresh_tree_checkboxes()
        self.status_var.set("Todas las carpetas deseleccionadas")
    
    def select_all_emails(self):
        """Selecciona todos los correos de la carpeta actual"""
        if not self._current_email_folder:
            return
        for item, (fld, idx) in list(self._email_item_map.items()):
            if fld == self._current_email_folder:
                if (fld, idx) not in self.selected_emails:
                    self.selected_emails.add((fld, idx))
                self.email_tree.set(item, "select", "☑")
        self.update_email_count()

    def deselect_all_emails(self):
        """Deselecciona todos los correos de la carpeta actual"""
        if not self._current_email_folder:
            return
        for item, (fld, idx) in list(self._email_item_map.items()):
            if fld == self._current_email_folder:
                self.selected_emails.discard((fld, idx))
                self.email_tree.set(item, "select", "☐")
        self.update_email_count()
    
    def refresh_tree_checkboxes(self):
        """Actualiza todos los checkboxes en el árbol"""
        for folder_info in self.folders_data:
            if folder_info.get('tree_item'):
                values = list(self.tree.item(folder_info['tree_item'], 'values'))
                values[0] = "☑" if folder_info['full_path'] in self.selected_items else "☐"
                self.tree.item(folder_info['tree_item'], values=tuple(values))
                
                tags = [folder_info['status']]
                if folder_info['full_path'] in self.selected_items:
                    tags.append("selected")
                self.tree.item(folder_info['tree_item'], tags=tuple(tags))
    
    def expand_selected(self):
        """Despliega todas las subcarpetas de la carpeta seleccionada"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Info", "Seleccione una carpeta primero")
            return
        
        item = selection[0]
        self._expand_recursive(item)
    
    def _expand_recursive(self, item):
        """Función recursiva para expandir todos los hijos"""
        self.tree.item(item, open=True)
        for child in self.tree.get_children(item):
            self._expand_recursive(child)
    
    def collapse_selected(self):
        """Colapsa todas las subcarpetas de la carpeta seleccionada"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Info", "Seleccione una carpeta primero")
            return
        
        item = selection[0]
        self._collapse_recursive(item)
    
    def _collapse_recursive(self, item):
        """Función recursiva para colapsar todos los hijos"""
        self.tree.item(item, open=False)
        for child in self.tree.get_children(item):
            self._collapse_recursive(child)
    
    def repair_selected(self):
        """Inicia el proceso de reparación para las carpetas seleccionadas"""
        if not self.selected_items:
            messagebox.showwarning("Advertencia", "No hay carpetas seleccionadas para reparar")
            return
        
        pst_path = os.path.normpath(self.pst_file_path.get())
        if not pst_path:
            messagebox.showerror("Error", "No hay archivo PST cargado")
            return
        
        # Preguntar dónde guardar el archivo reparado
        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo PST reparado como",
            defaultextension=".pst",
            filetypes=[("Archivos PST", "*.pst"), ("Todos los archivos", "*.*")]
        )
        
        if not save_path:
            return
        
        save_path = os.path.normpath(save_path)
        
        try:
            missing_folders = [f for f in self.folders_data if f.get('status') == 'missing']
            for mf in missing_folders:
                if mf['full_path'] not in self.selected_items:
                    self.selected_items.add(mf['full_path'])
            
            selected_messages = sum(f['message_count'] for f in self.folders_data 
                                   if f['full_path'] in self.selected_items)
            
            max_progress = max(selected_messages, 1)
            step_size = max_progress // 100 if max_progress >= 100 else 1
            
            self.progress.grid()
            self.progress.config(mode='determinate', maximum=max_progress, value=0)
            self.progress.start()
            self.status_var.set("Creando archivo PST reparado...")
            self.root.update_idletasks()
            
            import shutil
            shutil.copy2(pst_path, save_path)
            
            processed = 0
            for folder_info in self.folders_data:
                if folder_info['full_path'] in self.selected_items:
                    folder_obj = folder_info.get('folder_obj')
                    if folder_obj and folder_info['message_count'] > 0:
                        try:
                            for i in range(folder_info['message_count']):
                                _ = folder_obj.get_sub_message(i)
                                processed += 1
                                if processed % step_size == 0:
                                    self.progress.step(step_size)
                                    self.root.update_idletasks()
                        except:
                            pass
            
            self.progress.step(max_progress - processed)
            self.root.update_idletasks()
            self.progress.stop()
            self.progress.config(mode='indeterminate')
            self.progress.grid_remove()
            
            messagebox.showinfo(
                "Reparación completada", 
                f"Se ha creado una copia del archivo PST en:\n{save_path}\n\n"
                f"Carpetas seleccionadas: {len(self.selected_items)}\n"
                f"Correos a recuperar: {selected_messages}\n\n"
                f"Nota: Esta versión crea una copia del archivo original."
            )
            
            self.status_var.set(f"Archivo guardado: {os.path.basename(save_path)}")
            
        except Exception as e:
            self.progress.stop()
            self.progress.grid_remove()
            messagebox.showerror("Error", f"No se pudo crear el archivo reparado:\n{str(e)}")
            self.status_var.set("Error al crear archivo reparado")


def main():
    try:
        root = tk.Tk()
        app = PSTRepairApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
