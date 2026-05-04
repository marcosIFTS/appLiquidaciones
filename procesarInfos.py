
import tkinter as tk
from tkinter import PhotoImage, filedialog, ttk, messagebox
import pandas as pd
import os
from datetime import date

try:
    from tkcalendar import DateEntry # type: ignore
except ImportError:
    DateEntry = None

try:
    from PIL import Image # type: ignore
    from PIL.ImageTk import PhotoImage as PILPhotoImage # type: ignore
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Configura los nombres de las columnas a consolidar
COLUMNAS_CONSOLIDADAS = ['claveIGT', 'Sucursal', 'cuit', 'nroDocumento', 'expediente', 'importe', 'carga']
COLUMNAS_CONSOLIDADO_CARGAS = ['carga', 'tipoIncentivo', 'exp', 'Idprograma', 'fechaPago', 'periodo', 'tipo', 'descripcion', 'casos', 'montoPago']
COLUMNAS_EDITABLES_PASO_2 = ['tipoIncentivo', 'fechaPago', 'periodo', 'tipo', 'descripcion']
PLACEHOLDER_PERIODO = 'aaaamm'
PLACEHOLDER_DESCRIPCION = 'breve descripcion de la carga'

class LiquidacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Liquidación - Procesamiento de Cargas")
        # Maximizar ventana para que la app ocupe toda la pantalla (Windows)
        try:
            self.root.state('zoomed')
        except Exception:
            self.root.geometry("1200x760")
        self.root.resizable(True, True)
        
        # Variables globales
        self.archivos_excel = []
        self.archivos_incorrectos = []
        self.datos_consolidados = []
        self.cargas_consolidadas = []
        self.df_consolidado = pd.DataFrame()
        self.df_consolidadoCargas = pd.DataFrame(columns=COLUMNAS_CONSOLIDADO_CARGAS)
        self.edicion_activa = None
        self.row_widgets = {}
        self.paso_actual = 1
        
        # Configurar estilos
        self.root.configure(bg="#f0f0f0")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Marco principal
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Asegurar que el contenido central pueda expandirse y empujar el footer hacia abajo
        self.main_frame.grid_rowconfigure(3, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        self.titulo = ttk.Label(self.main_frame, text="Sistema de Liquidación", font=("Arial", 18, "bold"))
        self.titulo.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Indicador de pasos
        self.pasos_frame = ttk.Frame(self.main_frame)
        self.pasos_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        self.pasos_labels = []
        for i in range(1, 4):
            label = ttk.Label(self.pasos_frame, text=f"Paso {i}", font=("Arial", 10))
            label.pack(side=tk.LEFT, padx=15)
            self.pasos_labels.append(label)
        
        # Separador
        ttk.Separator(self.main_frame, orient="horizontal").grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Marco de contenido dinámico
        self.contenido_frame = ttk.Frame(self.main_frame)
        self.contenido_frame.grid(row=3, column=0, columnspan=2, pady=20, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Marco de botones
        self.botones_frame = ttk.Frame(self.main_frame)
        self.botones_frame.grid(row=4, column=0, columnspan=2, pady=20, sticky=(tk.E, tk.W))
        
        # Footer fijo en la ventana principal para que sea visible desde el inicio
        self.footer_frame = ttk.Frame(self.root, padding=(20, 4, 20, 10))
        # Ubicar el footer en la fila inferior y anclarlo al sur para que permanezca abajo
        self.footer_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.S))
        self.root.grid_rowconfigure(1, weight=0)
        self.footer_frame.grid_columnconfigure(0, weight=1)

        self.footer_inner = ttk.Frame(self.footer_frame)
        self.footer_inner.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Cargar logo con escalado responsivo
        self.dir_img = os.path.join(os.path.dirname(__file__), 'img', 'datosPAS_pie.png')
        self.logo_img = None
        self.logo_label = None
        self.img_original_pil = None  # Almacenar imagen original para rescalados
        self._cargar_y_escalar_logo_footer()
        
        # Bindear evento de redimensionamiento para rescalar dinámicamente
        self.root.bind('<Configure>', self._on_window_resize)
        
        self.mostrar_paso_1()

    def _reset_progreso_paso1(self, mensaje=''):
        if hasattr(self, 'progreso_paso1'):
            self.progreso_paso1.configure(value=0, maximum=100)
        if hasattr(self, 'label_progreso_paso1'):
            self.label_progreso_paso1.configure(text=mensaje)

    def _actualizar_progreso_paso1(self, actual, total, mensaje=''):
        if total <= 0:
            total = 1
        porcentaje = (actual / total) * 100
        self.progreso_paso1.configure(value=porcentaje)
        if mensaje:
            self.label_progreso_paso1.configure(text=mensaje)
        self.root.update_idletasks()
    
    def _cargar_y_escalar_logo_footer(self):
        """Carga y escala la imagen del logo responsivamente"""
        # Limpiar logo anterior si existe
        if self.logo_label is not None:
            self.logo_label.destroy()
            self.logo_label = None
            self.logo_img = None
        
        if not os.path.exists(self.dir_img):
            self.logo_label = ttk.Label(self.footer_inner, text="Logo no encontrado")
            self.logo_label.pack(side=tk.TOP, pady=4, anchor=tk.CENTER)
            return
        
        try:
            # Altura máxima del footer
            altura_maxima = 90
            
            # Cargar imagen original solo una vez
            if self.img_original_pil is None and HAS_PIL:
                self.img_original_pil = Image.open(self.dir_img)
            
            # Calcular ancho disponible del footer
            self.root.update_idletasks()  # Asegurar que el layout esté actualizado
            ancho_disponible = self.root.winfo_width()   # Restar padding y márgenes

            if HAS_PIL and self.img_original_pil is not None:             
                # Redimensionar imagen
                img_redimensionada = self.img_original_pil.resize((1200, 760), Image.Resampling.LANCZOS)
                self.logo_img = PILPhotoImage(img_redimensionada)
            else:
                # Fallback: usar PhotoImage con subsample si PIL no está disponible
                if self.logo_img is None:
                    self.logo_img = PhotoImage(file=self.dir_img)
                
                # Obtener dimensiones originales
                ancho_original = self.logo_img.width()
                alto_original = self.logo_img.height()
                
                # Calcular factor de reducción basado en ancho disponible
                factor = max(1, max(ancho_original // ancho_disponible, alto_original // altura_maxima))
                self.logo_img = self.logo_img.subsample(factor, factor)
                
            self.logo_label = ttk.Label(self.footer_inner, image=self.logo_img)
            self.logo_label.pack(side=tk.TOP, pady=4, anchor=tk.CENTER)
        except Exception as e:
            print(f"Error al cargar logo: {e}")
            self.logo_label = ttk.Label(self.footer_inner, text="Logo no disponible")
            self.logo_label.pack(side=tk.TOP, pady=4, anchor=tk.CENTER)
    
    def _on_window_resize(self, event):
        """Maneja el evento de redimensionamiento de ventana para rescalar logo"""
        if event.widget == self.root and HAS_PIL:
            # Rescalar imagen cuando la ventana se redimensiona significativamente
            self.root.after(500, self._cargar_y_escalar_logo_footer)
        
    def limpiar_contenido(self):
        """Limpia el marco de contenido"""
        for widget in self.contenido_frame.winfo_children():
            widget.destroy()
        for widget in self.botones_frame.winfo_children():
            widget.destroy()
            
    def actualizar_indicador_pasos(self):
        """Actualiza el color del indicador de pasos"""
        for i, label in enumerate(self.pasos_labels, 1):
            if i == self.paso_actual:
                label.configure(foreground="#007acc", font=("Arial", 10, "bold"))
            elif i < self.paso_actual:
                label.configure(foreground="#28a745", font=("Arial", 10))
            else:
                label.configure(foreground="#6c757d", font=("Arial", 10))
    
    def validar_archivo_excel(self, ruta_archivo):
        """
        Valida que un archivo Excel tenga las columnas requeridas
        Retorna: (es_valido, mensaje)
        """
        try:
            # Leer solo los encabezados
            df = pd.read_excel(ruta_archivo, nrows=0)
            columnas_archivo = set(df.columns.tolist())
            columnas_requeridas = set(COLUMNAS_CONSOLIDADAS)
            
            # Verificar que tenga todas las columnas
            columnas_faltantes = columnas_requeridas - columnas_archivo
            
            if columnas_faltantes:
                msg = f"Faltan columnas: {', '.join(columnas_faltantes)}"
                return False, msg
            
            return True, "OK"
        except Exception as e:
            return False, f"Error al leer: {str(e)}"

    def _es_archivo_duplicado(self, ruta_archivo):
        nombre_archivo = os.path.basename(ruta_archivo)
        return any(os.path.basename(archivo) == nombre_archivo for archivo in self.archivos_excel) or any(
            os.path.basename(archivo_incorrecto[0]) == nombre_archivo for archivo_incorrecto in self.archivos_incorrectos
        )

    def _procesar_archivos_seleccionados(self, archivos, agregar=False):
        if not agregar:
            self.archivos_excel = []
            self.archivos_incorrectos = []
            self.tree_paso1.delete(*self.tree_paso1.get_children())

        self._reset_progreso_paso1("Validando archivos...")
        nuevos_invalidos = []
        total_archivos = len(archivos)
        procesados = 0

        for archivo in archivos:
            procesados += 1
            if self._es_archivo_duplicado(archivo):
                self._actualizar_progreso_paso1(procesados, total_archivos, f"Validando archivos... {procesados}/{total_archivos}")
                continue

            nombre_archivo = os.path.basename(archivo)
            es_valido, mensaje = self.validar_archivo_excel(archivo)

            if es_valido:
                self.archivos_excel.append(archivo)
                estado = "Válido"
                tag = "valido"
            else:
                self.archivos_incorrectos.append((archivo, mensaje))
                nuevos_invalidos.append((nombre_archivo, mensaje))
                estado = f"Incorrecto: {mensaje}"
                tag = "invalido"

            self.tree_paso1.insert("", "end", values=(nombre_archivo, estado), tags=(tag,))
            self._actualizar_progreso_paso1(procesados, total_archivos, f"Validando archivos... {procesados}/{total_archivos}")

        self.tree_paso1.tag_configure("valido", foreground="green")
        self.tree_paso1.tag_configure("invalido", foreground="red")
        self._actualizar_resumen_paso1()

        estado_habilitado = tk.NORMAL if self.archivos_excel else tk.DISABLED
        self.btn_agregar.configure(state=estado_habilitado)
        self.btn_confirmar_paso1.configure(state=estado_habilitado)

        if total_archivos > 0:
            self._actualizar_progreso_paso1(total_archivos, total_archivos, "Validación completada.")

        if nuevos_invalidos:
            detalles = "\n".join([f"  - {nombre}: {mensaje}" for nombre, mensaje in nuevos_invalidos])
            messagebox.showwarning("Archivos incorrectos", "Estos archivos no respetan el formato:\n" + detalles)
    
    def mostrar_paso_1(self):
        """Muestra la interfaz del Paso 1: Selección de archivos Excel"""
        self.limpiar_contenido()
        self.paso_actual = 1
        self.actualizar_indicador_pasos()
        
        # Título del paso
        titulo = ttk.Label(self.contenido_frame, text="Paso 1: Seleccionar Archivos Excel", font=("Arial", 14, "bold"))
        titulo.pack(pady=10)
        
        # Instrucciones
        instrucciones = ttk.Label(
            self.contenido_frame, 
            text="Selecciona archivos Excel que contengan las siguientes columnas:\n" + 
                 ", ".join(COLUMNAS_CONSOLIDADAS),
            font=("Arial", 10),
            justify=tk.CENTER
        )
        instrucciones.pack(pady=10)
        
        # Botón para seleccionar archivos
        btn_seleccionar = ttk.Button(
            self.contenido_frame,
            text="📁 Seleccionar Archivos Excel",
            command=self.seleccionar_archivos
        )
        btn_seleccionar.pack(pady=10)
        
        # Botón para agregar más archivos
        btn_agregar = ttk.Button(
            self.contenido_frame,
            text="➕ Agregar Más Archivos",
            command=self.agregar_mas_archivos,
            state=tk.DISABLED
        )
        btn_agregar.pack(pady=5)
        self.btn_agregar = btn_agregar

        btn_confirmar = ttk.Button(
            self.contenido_frame,
            text="Confirmar archivos y pasar al Paso 2",
            command=self.ir_paso_2,
            state=tk.DISABLED
        )
        btn_confirmar.pack(pady=5)
        self.btn_confirmar_paso1 = btn_confirmar
        
        # Marco para mostrar archivos seleccionados
        marco_archivos = ttk.LabelFrame(self.contenido_frame, text="Archivos Seleccionados", padding="10")
        marco_archivos.pack(pady=10, fill=tk.BOTH, expand=True)
        
        # Treeview para mostrar archivos
        columnas_tree = ("Archivo", "Estado")
        self.tree_paso1 = ttk.Treeview(marco_archivos, columns=columnas_tree, height=12, show="headings")
        
        self.tree_paso1.heading("Archivo", text="Nombre del Archivo")
        self.tree_paso1.heading("Estado", text="Estado")
        
        self.tree_paso1.column("Archivo", width=600)
        self.tree_paso1.column("Estado", width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(marco_archivos, orient=tk.VERTICAL, command=self.tree_paso1.yview)
        self.tree_paso1.configure(yscroll=scrollbar.set)
        
        self.tree_paso1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Etiqueta de resumen
        self.label_resumen_paso1 = ttk.Label(self.contenido_frame, text="", font=("Arial", 9), foreground="blue")
        self.label_resumen_paso1.pack(pady=5)

        # Barra de progreso para validacion y consolidacion
        self.label_progreso_paso1 = ttk.Label(self.contenido_frame, text="", font=("Arial", 9), foreground="#555555")
        self.label_progreso_paso1.pack(pady=(3, 2))

        self.progreso_paso1 = ttk.Progressbar(self.contenido_frame, orient='horizontal', mode='determinate', length=560)
        self.progreso_paso1.pack(pady=(0, 8), fill=tk.X, padx=20)
        self._reset_progreso_paso1("Esperando seleccion de archivos...")
        
        # Botones de acción
        btn_limpiar = ttk.Button(
            self.botones_frame,
            text="🗑️ Limpiar Lista",
            command=self.limpiar_paso_1
        )
        btn_limpiar.pack(side=tk.RIGHT, padx=5)
    
    def seleccionar_archivos(self):
        """Abre diálogo para seleccionar archivos Excel"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        
        if archivos:
            self._procesar_archivos_seleccionados(archivos)
    
    def agregar_mas_archivos(self):
        """Permite agregar más archivos a la selección existente"""
        archivos = filedialog.askopenfilenames(
            title="Agregar más archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        
        if archivos:
            self._procesar_archivos_seleccionados(archivos, agregar=True)
    
    def _actualizar_resumen_paso1(self):
        """Actualiza el resumen de archivos en Paso 1"""
        resumen = f"Válidos: {len(self.archivos_excel)} | Inválidos: {len(self.archivos_incorrectos)}"
        self.label_resumen_paso1.configure(text=resumen)
        
        if self.archivos_incorrectos:
            detalles = "\n".join([f"  • {os.path.basename(f[0])}: {f[1]}" for f in self.archivos_incorrectos])
            messagebox.showinfo("Archivos Inválidos", "Los siguientes archivos tienen problemas:\n" + detalles)
    
    def limpiar_paso_1(self):
        """Limpia la selección de archivos en Paso 1"""
        self.archivos_excel = []
        self.archivos_incorrectos = []
        self.tree_paso1.delete(*self.tree_paso1.get_children())
        self.label_resumen_paso1.configure(text="")
        self.btn_agregar.configure(state=tk.DISABLED)
        self.btn_confirmar_paso1.configure(state=tk.DISABLED)
    
    def ir_paso_2(self):
        """Confirma archivos y construye el consolidado del Paso 2"""
        if not self.archivos_excel:
            messagebox.showwarning("Paso 1 Incompleto", "Debes seleccionar al menos un archivo válido.")
            return

        self.datos_consolidados = []
        self.cargas_consolidadas = []

        # Deshabilita botones mientras consolida
        self.btn_confirmar_paso1.configure(state=tk.DISABLED)
        self.btn_agregar.configure(state=tk.DISABLED)
        self._reset_progreso_paso1("Consolidando archivos...")

        total_archivos = len(self.archivos_excel)
        procesados = 0

        for archivo in self.archivos_excel:
            df = pd.read_excel(archivo, usecols=COLUMNAS_CONSOLIDADAS)
            self.datos_consolidados.append(df)
            procesados += 1
            self._actualizar_progreso_paso1(procesados, total_archivos, f"Consolidando archivos... {procesados}/{total_archivos}")

        self._actualizar_progreso_paso1(total_archivos, total_archivos, "Generando datasets consolidados...")
        self.df_consolidado = pd.concat(self.datos_consolidados, ignore_index=True)
        self.df_consolidado[['claveIGT', 'Sucursal', 'cuit','nroDocumento','importe','carga']].to_csv('tmp_procCargasReg.txt', sep='|', index=False)

        reporte_agrupado = self.df_consolidado.groupby(['carga','expediente']).agg(
                Suma_monto=('importe', 'sum'),
                Cantidad=('cuit', 'count')
            )

        for idx, row in reporte_agrupado.iterrows():
            carga = [f"{idx[0]}",'SUBSIDIO',f"{idx[1]}","6",'dd/mm/aaaa','aaaamm','REGULAR/RETROACTIVA','breve descripcion carga',f"{row['Cantidad']}",f"{row['Suma_monto']}"]
            self.cargas_consolidadas.append(carga)

        self.df_consolidadoCargas = pd.DataFrame(self.cargas_consolidadas,columns=COLUMNAS_CONSOLIDADO_CARGAS)
        self._actualizar_progreso_paso1(100, 100, "Consolidacion completada. Abriendo Paso 2...")

        self.mostrar_paso_2()

    def mostrar_paso_2(self):
        """Muestra el Paso 2 con el consolidado editable"""
        self.limpiar_contenido()
        self.paso_actual = 2
        self.actualizar_indicador_pasos()

        titulo = ttk.Label(self.contenido_frame, text="Paso 2: Consolidado editable", font=("Arial", 14, "bold"))
        titulo.pack(pady=10)
        
        info = ttk.Label(
            self.contenido_frame,
            text=f"Se consolidaron {len(self.df_consolidadoCargas)} filas a partir de {len(self.archivos_excel)} archivo(s)",
            font=("Arial", 11)
        )
        info.pack(pady=8)

        marco_tabla = ttk.LabelFrame(self.contenido_frame, text="df_consolidadoCargas", padding="10")
        marco_tabla.pack(fill=tk.X, expand=False, padx=5, pady=6)

        self.tree_paso2 = ttk.Treeview(marco_tabla, columns=COLUMNAS_CONSOLIDADO_CARGAS, height=7, show="headings")
        for columna in COLUMNAS_CONSOLIDADO_CARGAS:
            self.tree_paso2.heading(columna, text=columna)
            ancho = 90
            if columna == 'descripcion':
                ancho = 240
            elif columna in ('carga', 'exp'):
                ancho = 170
            elif columna in COLUMNAS_EDITABLES_PASO_2:
                ancho = 130
            self.tree_paso2.column(columna, width=ancho, anchor=tk.W)

        scrollbar_y = ttk.Scrollbar(marco_tabla, orient=tk.VERTICAL, command=self.tree_paso2.yview)
        scrollbar_x = ttk.Scrollbar(marco_tabla, orient=tk.HORIZONTAL, command=self.tree_paso2.xview)
        self.tree_paso2.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_paso2.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree_paso2.bind('<Double-1>', self._iniciar_edicion_paso2)
        self._cargar_tree_paso2()

        marco_edicion = ttk.LabelFrame(self.contenido_frame, text="Editar fila seleccionada", padding="10")
        marco_edicion.pack(fill=tk.X, padx=5, pady=4)
        self._crear_formulario_edicion(marco_edicion)

        self.label_estado_paso2 = ttk.Label(self.contenido_frame, text="Doble clic en una fila para editar las columnas permitidas.", foreground="blue")
        self.label_estado_paso2.pack(pady=5)
        
        # Botones
        btn_volver = ttk.Button(
            self.botones_frame,
            text="⬅ Volver al Paso 1",
            command=self.mostrar_paso_1
        )
        btn_volver.pack(side=tk.LEFT, padx=5)
        
        btn_continuar = ttk.Button(
            self.botones_frame,
            text="Confirmar cambios y pasar al Paso 3 ➜",
            command=self.confirmar_cambios_paso2
        )
        btn_continuar.pack(side=tk.RIGHT, padx=5)

    def confirmar_cambios_paso2(self):
        if self.edicion_activa is not None:
            self.aplicar_edicion_paso2()

        self.ir_paso_3()

    def _cargar_tree_paso2(self):
        self.tree_paso2.delete(*self.tree_paso2.get_children())
        self.row_widgets = {}
        for indice, fila in self.df_consolidadoCargas.iterrows():
            valores = [fila[columna] for columna in COLUMNAS_CONSOLIDADO_CARGAS]
            item_id = self.tree_paso2.insert("", "end", values=valores)
            self.row_widgets[item_id] = indice

    def _crear_formulario_edicion(self, parent):
        self.edicion_vars = {}
        self.edicion_widgets = {}

        # tipoIncentivo: selector
        ttk.Label(parent, text='tipoIncentivo').grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        self.edicion_vars['tipoIncentivo'] = tk.StringVar(value='SUBSIDIO')
        combo_tipo_incentivo = ttk.Combobox(
            parent,
            textvariable=self.edicion_vars['tipoIncentivo'],
            values=['SUBSIDIO', 'EXTRAORDINARIA'],
            state='readonly',
            width=42,
        )
        combo_tipo_incentivo.grid(row=0, column=1, sticky=tk.W, padx=4, pady=4)
        self.edicion_widgets['tipoIncentivo'] = combo_tipo_incentivo

        # fechaPago: calendario con fecha de hoy
        ttk.Label(parent, text='fechaPago').grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        self.edicion_vars['fechaPago'] = tk.StringVar(value=date.today().strftime('%d/%m/%Y'))
        if DateEntry is not None:
            fecha_widget = DateEntry(
                parent,
                textvariable=self.edicion_vars['fechaPago'],
                date_pattern='dd/mm/yyyy',
                width=42,
            )
        else:
            fecha_widget = ttk.Entry(parent, textvariable=self.edicion_vars['fechaPago'], width=45)
        fecha_widget.grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)
        self.edicion_widgets['fechaPago'] = fecha_widget

        # periodo: placeholder aaaamm
        ttk.Label(parent, text='periodo').grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
        self.edicion_vars['periodo'] = tk.StringVar()
        entry_periodo = ttk.Entry(parent, textvariable=self.edicion_vars['periodo'], width=45)
        entry_periodo.grid(row=2, column=1, sticky=tk.W, padx=4, pady=4)
        self._set_placeholder(entry_periodo, self.edicion_vars['periodo'], PLACEHOLDER_PERIODO)
        self.edicion_widgets['periodo'] = entry_periodo

        # tipo: selector
        ttk.Label(parent, text='tipo').grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
        self.edicion_vars['tipo'] = tk.StringVar(value='REGULAR')
        combo_tipo = ttk.Combobox(
            parent,
            textvariable=self.edicion_vars['tipo'],
            values=['REGULAR', 'RETROACTIVA'],
            state='readonly',
            width=42,
        )
        combo_tipo.grid(row=3, column=1, sticky=tk.W, padx=4, pady=4)
        self.edicion_widgets['tipo'] = combo_tipo

        # descripcion: placeholder breve descripcion de la carga
        ttk.Label(parent, text='descripcion').grid(row=4, column=0, sticky=tk.W, padx=4, pady=4)
        self.edicion_vars['descripcion'] = tk.StringVar()
        entry_descripcion = ttk.Entry(parent, textvariable=self.edicion_vars['descripcion'], width=45)
        entry_descripcion.grid(row=4, column=1, sticky=tk.W, padx=4, pady=4)
        self._set_placeholder(entry_descripcion, self.edicion_vars['descripcion'], PLACEHOLDER_DESCRIPCION)
        self.edicion_widgets['descripcion'] = entry_descripcion

        botones = ttk.Frame(parent)
        botones.grid(row=0, column=2, rowspan=len(COLUMNAS_EDITABLES_PASO_2), padx=10, sticky=tk.N)

        ttk.Button(botones, text="Aplicar edición", command=self.aplicar_edicion_paso2).pack(fill=tk.X, pady=2)
        ttk.Button(botones, text="Cancelar", command=self.cancelar_edicion_paso2).pack(fill=tk.X, pady=2)

    def _set_placeholder(self, entry_widget, variable, placeholder_text):
        variable.set(placeholder_text)
        entry_widget.configure(foreground='gray')

        def on_focus_in(_event):
            if variable.get() == placeholder_text:
                variable.set('')
                entry_widget.configure(foreground='black')

        def on_focus_out(_event):
            if not variable.get().strip():
                variable.set(placeholder_text)
                entry_widget.configure(foreground='gray')

        entry_widget.bind('<FocusIn>', on_focus_in)
        entry_widget.bind('<FocusOut>', on_focus_out)

    def _set_entry_value(self, campo, valor):
        widget = self.edicion_widgets.get(campo)
        variable = self.edicion_vars.get(campo)
        if widget is None or variable is None:
            return

        if campo == 'periodo':
            texto = str(valor).strip() if str(valor).strip() else PLACEHOLDER_PERIODO
            variable.set(texto)
            widget.configure(foreground='black' if texto != PLACEHOLDER_PERIODO else 'gray')
            return

        if campo == 'descripcion':
            texto = str(valor).strip() if str(valor).strip() else PLACEHOLDER_DESCRIPCION
            variable.set(texto)
            widget.configure(foreground='black' if texto != PLACEHOLDER_DESCRIPCION else 'gray')
            return

        variable.set(str(valor))

    def _get_valor_edicion(self, campo):
        valor = self.edicion_vars[campo].get().strip()
        if campo == 'periodo' and valor == PLACEHOLDER_PERIODO:
            return ''
        if campo == 'descripcion' and valor == PLACEHOLDER_DESCRIPCION:
            return ''
        return valor

    def _iniciar_edicion_paso2(self, event):
        item = self.tree_paso2.identify_row(event.y)
        if not item:
            return

        self.edicion_activa = item
        indice = self.row_widgets[item]
        fila = self.df_consolidadoCargas.iloc[indice]

        for columna in COLUMNAS_EDITABLES_PASO_2:
            self._set_entry_value(columna, fila[columna])

        self.label_estado_paso2.configure(text=f"Editando fila con carga {fila['carga']} y expediente {fila['exp']}")

    def aplicar_edicion_paso2(self):
        if self.edicion_activa is None:
            messagebox.showinfo("Sin seleccion", "Selecciona una fila con doble clic antes de editar.")
            return

        indice = self.row_widgets[self.edicion_activa]
        for columna in COLUMNAS_EDITABLES_PASO_2:
            self.df_consolidadoCargas.at[indice, columna] = self._get_valor_edicion(columna)

        for columna in COLUMNAS_CONSOLIDADO_CARGAS:
            self.tree_paso2.set(self.edicion_activa, columna, self.df_consolidadoCargas.at[indice, columna])

        self.label_estado_paso2.configure(text="Cambios aplicados en la fila seleccionada.")

    def cancelar_edicion_paso2(self):
        self.edicion_activa = None
        self.edicion_vars['tipoIncentivo'].set('SUBSIDIO')
        self.edicion_vars['tipo'].set('REGULAR')
        self.edicion_vars['fechaPago'].set(date.today().strftime('%d/%m/%Y'))
        self._set_entry_value('periodo', '')
        self._set_entry_value('descripcion', '')
        self.label_estado_paso2.configure(text="Edicion cancelada.")
    
    def ir_paso_3(self):
        """Continúa al Paso 3"""
        self.mostrar_paso_3()
    
    def mostrar_paso_3(self):
        """Muestra el resumen de datasets y permite exportarlos"""
        self.limpiar_contenido()
        self.paso_actual = 3
        self.actualizar_indicador_pasos()
        
        titulo = ttk.Label(self.contenido_frame, text="Paso 3: Resumen y exportación", font=("Arial", 14, "bold"))
        titulo.pack(pady=10)
        
        marco_resumen = ttk.LabelFrame(self.contenido_frame, text="Resumen de registros", padding="10")
        marco_resumen.pack(fill=tk.X, padx=5, pady=10)

        resumen_texto = (
            f"Archivos Excel válidos: {len(self.archivos_excel)}\n"
            f"Archivos incorrectos: {len(self.archivos_incorrectos)}\n"
            f"Registros en tmp_procCargasReg: {len(self.df_consolidado)}\n"
            f"Registros en cargas_consolidadas: {len(self.df_consolidadoCargas)}"
        )
        ttk.Label(marco_resumen, text=resumen_texto, justify=tk.LEFT, font=("Arial", 11)).pack(anchor=tk.W)

        marco_exportacion = ttk.LabelFrame(self.contenido_frame, text="Exportación", padding="10")
        marco_exportacion.pack(fill=tk.X, padx=5, pady=10)

        ttk.Label(
            marco_exportacion,
            text="Exporta los datasets consolidados a archivos TXT con separador '|'.",
            justify=tk.LEFT,
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=(0, 8))

        ttk.Button(
            marco_exportacion,
            text="Exportar datasets a TXT",
            command=self.exportar_datasets_csv
        ).pack(anchor=tk.W)
        
        # Botones
        btn_volver = ttk.Button(
            self.botones_frame,
            text="⬅ Volver",
            command=self.mostrar_paso_2
        )
        btn_volver.pack(side=tk.LEFT, padx=5)
        
        btn_finalizar = ttk.Button(
            self.botones_frame,
            text="✓ Finalizar",
            command=self.finalizar
        )
        btn_finalizar.pack(side=tk.RIGHT, padx=5)

    def exportar_datasets_csv(self):
        ruta_base = os.getcwd()
        ruta_datos = os.path.join(ruta_base, 'tmp_procCargasReg.txt')
        ruta_cargas = os.path.join(ruta_base, 'tmp_procCargas.txt')

        self.df_consolidado.to_csv(ruta_datos, sep='|', index=False)
        self.df_consolidadoCargas.to_csv(ruta_cargas, sep='|', index=False)
        messagebox.showinfo(
            "Exportación completada",
            f"Se exportaron los archivos en:\n- {ruta_datos}\n- {ruta_cargas}"
        )
    
    def finalizar(self):
        """Finaliza el proceso"""
        messagebox.showinfo("Completado", "Proceso finalizado correctamente.")
        self.root.quit()


if __name__ == "__main__":
    root = tk.Tk()
    app = LiquidacionApp(root)
    root.mainloop()
