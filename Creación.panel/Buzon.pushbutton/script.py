# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
import csv
import clr
import math
import os

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")

from System import String
from System.Data import DataTable
from System.Windows import Visibility, WindowStartupLocation
from System.Windows.Controls import DataGridTextColumn
from System.Windows.Data import Binding
from System.Windows.Input import Key
from Microsoft.Win32 import OpenFileDialog

doc = revit.doc
XAML_FILE = os.path.join(os.path.dirname(__file__), "ui.xaml")


# =========================================================
# UTILIDADES
# =========================================================
def metros_a_interno(valor_m):
    try:
        return DB.UnitUtils.ConvertToInternalUnits(valor_m, DB.UnitTypeId.Meters)
    except:
        return DB.UnitUtils.ConvertToInternalUnits(valor_m, DB.DisplayUnitType.DUT_METERS)


def obtener_nivel_mas_cercano(niveles, elevacion_obj):
    if not niveles:
        return None
    return min(niveles, key=lambda lv: abs(lv.Elevation - elevacion_obj))


def get_all_levels(doc):
    niveles = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    return sorted(niveles, key=lambda x: x.Elevation)


def safe_str(x):
    if x is None:
        return ""
    return str(x)


def try_parse_float(texto):
    try:
        return float(safe_str(texto).strip().replace(",", "."))
    except:
        return None


def get_symbol_display_name(symbol):
    fam_name = ""
    typ_name = ""

    try:
        fam_name = symbol.Family.Name
    except:
        fam_name = ""

    try:
        typ_name = DB.Element.Name.GetValue(symbol)
    except:
        try:
            typ_name = symbol.Name
        except:
            typ_name = ""

    return fam_name, typ_name


def get_insertable_categories(doc):
    symbols = list(DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol))
    cat_map = {}

    for s in symbols:
        try:
            cat = s.Category
            if cat and cat.Name:
                cat_id = get_element_id_int(cat.Id)
                if cat_id not in cat_map:
                    cat_map[cat_id] = cat
        except:
            pass

    cats = sorted(cat_map.values(), key=lambda c: c.Name)
    return cats


def get_element_id_int(element_id):
    try:
        return element_id.IntegerValue
    except:
        return element_id.Value


def get_symbols_by_category(doc, category_id_int):
    symbols = list(DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol))
    result = []

    for s in symbols:
        try:
            if s.Category and get_element_id_int(s.Category.Id) == category_id_int:
                result.append(s)
        except:
            pass

    return sorted(
        result,
        key=lambda x: (
            safe_str(x.Family.Name),
            safe_str(DB.Element.Name.GetValue(x))
        )
    )


def agrupar_simbolos(doc):
    data = {}
    categorias = get_insertable_categories(doc)

    for cat in categorias:
        data[cat.Name] = {}
        symbols = get_symbols_by_category(doc, get_element_id_int(cat.Id))

        for s in symbols:
            fam_name, typ_name = get_symbol_display_name(s)

            if fam_name not in data[cat.Name]:
                data[cat.Name][fam_name] = []

            data[cat.Name][fam_name].append(s)

    return data


def buscar_symbol_por_texto(data, categoria, familia, tipo):
    if categoria not in data:
        return None

    if familia not in data[categoria]:
        return None

    for s in data[categoria][familia]:
        fam_name, typ_name = get_symbol_display_name(s)
        if typ_name == tipo:
            return s

    return None


def activar_symbol_si_es_necesario(symbol):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()


def crear_instancia_generica(doc, symbol, punto, nivel):
    activar_symbol_si_es_necesario(symbol)

    try:
        return doc.Create.NewFamilyInstance(
            punto,
            symbol,
            nivel,
            DB.Structure.StructuralType.NonStructural
        )
    except:
        pass

    try:
        return doc.Create.NewFamilyInstance(
            punto,
            symbol,
            DB.Structure.StructuralType.NonStructural
        )
    except:
        pass

    try:
        return doc.Create.NewFamilyInstance(
            punto,
            symbol,
            nivel
        )
    except:
        pass

    raise Exception("No se pudo crear la instancia para la familia/tipo seleccionado.")


def asignar_comentario(elem, texto):
    p = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p and not p.IsReadOnly:
        p.Set(texto if texto else "")


def asignar_parametro_texto(elem, nombre_parametro, valor):
    try:
        p = elem.LookupParameter(nombre_parametro)
        if p and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
            p.Set(valor if valor else "")
            return True
    except:
        pass

    return False


def asignar_nivel_si_aplica(elem, nivel):
    if not nivel:
        return

    posibles = [
        DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
        DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM
    ]

    for bip in posibles:
        try:
            p = elem.get_Parameter(bip)
            if p and not p.IsReadOnly and p.StorageType == DB.StorageType.ElementId:
                p.Set(nivel.Id)
                return
        except:
            pass


# =========================================================
# COORDENADAS
# CSV = Shared Coordinates exactas
# =========================================================
def shared_a_internal_xyz(doc, norte_m, este_m, cota_m):
    project_position = doc.ActiveProjectLocation.GetProjectPosition(DB.XYZ.Zero)

    este_shared = metros_a_interno(este_m)
    norte_shared = metros_a_interno(norte_m)
    elev_shared = metros_a_interno(cota_m)

    dx = este_shared - project_position.EastWest
    dy = norte_shared - project_position.NorthSouth
    dz = elev_shared - project_position.Elevation

    ang = project_position.Angle
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)

    x = (dx * cos_a) - (dy * sin_a)
    y = (dx * sin_a) + (dy * cos_a)
    z = dz

    return DB.XYZ(x, y, z)


# =========================================================
# LECTURA CSV
# =========================================================
def leer_csv_puntos(ruta_csv, categoria_default, familia_default, tipo_default, usar_columnas_syp=True):
    resultado = []

    with open(ruta_csv, 'r') as f:
        sample = f.read(2048)
        f.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=';,')
        except:
            dialect = csv.excel
            dialect.delimiter = ','

        reader = csv.reader(f, dialect)
        filas = list(reader)

        if len(filas) <= 1:
            raise Exception("El CSV no tiene filas validas.")

        for i, row in enumerate(filas[1:], start=2):
            if len(row) < 4:
                continue

            comentario = safe_str(row[0]).strip()
            norte = try_parse_float(row[1])
            este = try_parse_float(row[2])
            cota = try_parse_float(row[3])

            cond_ingreso = ""
            cond_salida = ""
            conexion = ""

            if usar_columnas_syp:
                if len(row) > 4:
                    cond_ingreso = safe_str(row[4]).strip()
                if len(row) > 5:
                    cond_salida = safe_str(row[5]).strip()
                if len(row) > 6:
                    conexion = safe_str(row[6]).strip()

            resultado.append({
                "fila": i,
                "comentario": comentario,
                "norte": norte,
                "este": este,
                "cota": cota,
                "categoria": categoria_default,
                "familia": familia_default,
                "tipo": tipo_default,
                "000_SYP_COND_INGRESO": cond_ingreso,
                "000_SYP_COND_SALIDA": cond_salida,
                "000_SYP_CONEXION": conexion
            })

    return resultado


def seleccionar_csv():
    dlg = OpenFileDialog()
    dlg.Title = "Seleccionar archivo CSV"
    dlg.Filter = "Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*"
    dlg.Multiselect = False
    if dlg.ShowDialog():
        return dlg.FileName
    return None


# =========================================================
# VENTANA WPF
# =========================================================
class ImportarCsvWindow(forms.WPFWindow):
    def __init__(self, catalogo):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.catalogo = catalogo
        self.usar_columnas_syp = None
        self.seleccion = None
        self.filas = []
        self.resultado = None
        self.cancelado = True
        self.review_table = None

        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.PreviewKeyDown += self.on_preview_key_down

        self.BtnModeSyp.Click += self.on_modo_con_syp
        self.BtnModeNoSyp.Click += self.on_modo_sin_syp
        self.BtnModeCancel.Click += self.on_cancel
        self.BtnExampleOk.Click += self.on_example_ok
        self.BtnExampleCancel.Click += self.on_cancel
        self.BtnSelectionOk.Click += self.on_selection_ok
        self.BtnSelectionCancel.Click += self.on_cancel
        self.BtnReviewCreate.Click += self.on_review_create
        self.BtnReviewCancel.Click += self.on_cancel
        self.BtnApplyTypeToRow.Click += self.on_apply_type_to_row

        self.CboCategoria.SelectionChanged += self.on_categoria_changed
        self.CboFamilia.SelectionChanged += self.on_familia_changed
        self.ReviewGrid.SelectionChanged += self.on_review_selection_changed
        self.RowCategoria.SelectionChanged += self.on_row_categoria_changed
        self.RowFamilia.SelectionChanged += self.on_row_familia_changed

        self.cargar_categorias()
        self.mostrar_paso("modo")

    def on_preview_key_down(self, sender, args):
        if args.Key == Key.Escape:
            args.Handled = True
            self.on_cancel(None, None)

    def mostrar_paso(self, paso):
        self.PanelModo.Visibility = Visibility.Collapsed
        self.PanelEjemplo.Visibility = Visibility.Collapsed
        self.PanelSeleccion.Visibility = Visibility.Collapsed
        self.PanelRevision.Visibility = Visibility.Collapsed

        self.BtnModeSyp.IsDefault = False
        self.BtnExampleOk.IsDefault = False
        self.BtnSelectionOk.IsDefault = False
        self.BtnReviewCreate.IsDefault = False

        if paso == "modo":
            self.PanelModo.Visibility = Visibility.Visible
            self.BtnModeSyp.IsDefault = True
            self.Title = "Modo de importacion CSV"
        elif paso == "ejemplo":
            self.PanelEjemplo.Visibility = Visibility.Visible
            self.BtnExampleOk.IsDefault = True
            self.Title = "Ejemplo de como debe venir el CSV"
        elif paso == "seleccion":
            self.PanelSeleccion.Visibility = Visibility.Visible
            self.BtnSelectionOk.IsDefault = True
            self.Title = "Seleccionar categoria, familia y tipo"
        elif paso == "revision":
            self.PanelRevision.Visibility = Visibility.Visible
            self.BtnReviewCreate.IsDefault = True
            self.Title = "Revisar y editar datos a crear"

    def on_modo_con_syp(self, sender, args):
        self.usar_columnas_syp = True
        self.cargar_ejemplo_csv()
        self.mostrar_paso("ejemplo")

    def on_modo_sin_syp(self, sender, args):
        self.usar_columnas_syp = False
        self.cargar_ejemplo_csv()
        self.mostrar_paso("ejemplo")

    def cargar_ejemplo_csv(self):
        table = DataTable("EjemploCSV")
        columnas = [
            ("comentario", "COMENTARIO"),
            ("norte", "NORTE (m)"),
            ("este", "ESTE (m)"),
            ("cota", "COTA (m)")
        ]

        if self.usar_columnas_syp:
            columnas.extend([
                ("cond_ingreso", "000_SYP_COND_INGRESO"),
                ("cond_salida", "000_SYP_COND_SALIDA"),
                ("conexion", "000_SYP_CONEXION")
            ])

        for nombre, encabezado in columnas:
            table.Columns.Add(encabezado, String)

        if self.usar_columnas_syp:
            datos = [
                ["TAPA_01", "100.0", "200.0", "0.0", "INGRESO_A", "SALIDA_A", "CONEXION_A"],
                ["TAPA_02", "120.0", "220.0", "0.0", "INGRESO_B", "SALIDA_B", "CONEXION_B"],
                ["TAPA_03", "150.0", "250.0", "0.0", "INGRESO_C", "SALIDA_C", "CONEXION_C"],
            ]
        else:
            datos = [
                ["TAPA_01", "100.0", "200.0", "0.0"],
                ["TAPA_02", "120.0", "220.0", "0.0"],
                ["TAPA_03", "150.0", "250.0", "0.0"],
            ]

        for fila in datos:
            row = table.NewRow()
            for idx, valor in enumerate(fila):
                row[idx] = valor
            table.Rows.Add(row)

        self.ExampleGrid.ItemsSource = table.DefaultView

    def on_example_ok(self, sender, args):
        self.mostrar_paso("seleccion")

    def cargar_categorias(self):
        cats = sorted(self.catalogo.keys())
        self.CboCategoria.ItemsSource = cats
        self.RowCategoria.ItemsSource = cats

        if len(cats) > 0:
            self.CboCategoria.SelectedIndex = 0

    def cargar_familias(self, categoria, combo_familia, combo_tipo):
        familias = []
        if categoria in self.catalogo:
            familias = sorted(self.catalogo[categoria].keys())

        combo_familia.ItemsSource = familias
        combo_tipo.ItemsSource = []

        if len(familias) > 0:
            combo_familia.SelectedIndex = 0

    def cargar_tipos(self, categoria, familia, combo_tipo):
        tipos = []
        if categoria in self.catalogo and familia in self.catalogo[categoria]:
            for s in self.catalogo[categoria][familia]:
                fam_name, typ_name = get_symbol_display_name(s)
                tipos.append(typ_name)

        combo_tipo.ItemsSource = tipos
        if len(tipos) > 0:
            combo_tipo.SelectedIndex = 0

    def on_categoria_changed(self, sender, args):
        categoria = safe_str(self.CboCategoria.SelectedItem)
        self.cargar_familias(categoria, self.CboFamilia, self.CboTipo)

    def on_familia_changed(self, sender, args):
        categoria = safe_str(self.CboCategoria.SelectedItem)
        familia = safe_str(self.CboFamilia.SelectedItem)
        self.cargar_tipos(categoria, familia, self.CboTipo)

    def on_selection_ok(self, sender, args):
        categoria = safe_str(self.CboCategoria.SelectedItem)
        familia = safe_str(self.CboFamilia.SelectedItem)
        tipo = safe_str(self.CboTipo.SelectedItem)

        if not categoria or not familia or not tipo:
            forms.alert("Debes seleccionar categoria, familia y tipo.")
            return

        self.seleccion = {
            "categoria": categoria,
            "familia": familia,
            "tipo": tipo
        }

        ruta_csv = seleccionar_csv()
        if not ruta_csv:
            forms.alert("No se selecciono archivo CSV.")
            return

        try:
            self.filas = leer_csv_puntos(
                ruta_csv,
                self.seleccion["categoria"],
                self.seleccion["familia"],
                self.seleccion["tipo"],
                self.usar_columnas_syp
            )
        except Exception as ex:
            forms.alert("Error leyendo CSV:\n{}".format(ex))
            return

        if not self.filas:
            forms.alert("No se encontraron filas validas en el CSV.")
            return

        self.cargar_revision()
        self.mostrar_paso("revision")

    def crear_columna_texto(self, nombre, encabezado, ancho, readonly=False):
        col = DataGridTextColumn()
        col.Header = encabezado
        col.Width = ancho
        col.IsReadOnly = readonly
        col.Binding = Binding("[{}]".format(nombre))
        self.ReviewGrid.Columns.Add(col)

    def cargar_revision(self):
        self.ReviewGrid.Columns.Clear()
        self.crear_columna_texto("fila", "FILA", 55, True)
        self.crear_columna_texto("comentario", "COMENTARIO", 130)
        self.crear_columna_texto("norte", "NORTE (m)", 95)
        self.crear_columna_texto("este", "ESTE (m)", 95)
        self.crear_columna_texto("cota", "COTA (m)", 80)
        self.crear_columna_texto("categoria", "CATEGORIA", 155)
        self.crear_columna_texto("familia", "FAMILIA", 170)
        self.crear_columna_texto("tipo", "TIPO", 170)

        if self.usar_columnas_syp:
            self.crear_columna_texto("000_SYP_COND_INGRESO", "000_SYP_COND_INGRESO", 180)
            self.crear_columna_texto("000_SYP_COND_SALIDA", "000_SYP_COND_SALIDA", 180)
            self.crear_columna_texto("000_SYP_CONEXION", "000_SYP_CONEXION", 160)
            self.TxtReviewHint.Text = "Puedes modificar comentario, coordenadas, cota, categoria, familia, tipo y parametros antes de crear."
        else:
            self.TxtReviewHint.Text = "Puedes modificar comentario, coordenadas, cota, categoria, familia y tipo antes de crear."

        table = DataTable("Revision")
        for nombre in [
            "fila",
            "comentario",
            "norte",
            "este",
            "cota",
            "categoria",
            "familia",
            "tipo",
            "000_SYP_COND_INGRESO",
            "000_SYP_COND_SALIDA",
            "000_SYP_CONEXION"
        ]:
            table.Columns.Add(nombre, String)

        for item in self.filas:
            row = table.NewRow()
            row["fila"] = safe_str(item["fila"])
            row["comentario"] = safe_str(item["comentario"])
            row["norte"] = safe_str(item["norte"])
            row["este"] = safe_str(item["este"])
            row["cota"] = safe_str(item["cota"])
            row["categoria"] = safe_str(item["categoria"])
            row["familia"] = safe_str(item["familia"])
            row["tipo"] = safe_str(item["tipo"])
            row["000_SYP_COND_INGRESO"] = safe_str(item.get("000_SYP_COND_INGRESO", ""))
            row["000_SYP_COND_SALIDA"] = safe_str(item.get("000_SYP_COND_SALIDA", ""))
            row["000_SYP_CONEXION"] = safe_str(item.get("000_SYP_CONEXION", ""))
            table.Rows.Add(row)

        self.review_table = table
        self.ReviewGrid.ItemsSource = table.DefaultView
        if self.ReviewGrid.Items.Count > 0:
            self.ReviewGrid.SelectedIndex = 0

    def get_selected_data_row(self):
        item = self.ReviewGrid.SelectedItem
        if item is None:
            return None
        try:
            return item.Row
        except:
            return None

    def on_review_selection_changed(self, sender, args):
        row = self.get_selected_data_row()
        if row is None:
            return

        categoria = safe_str(row["categoria"])
        familia = safe_str(row["familia"])
        tipo = safe_str(row["tipo"])

        self.RowCategoria.SelectedItem = categoria
        if self.RowCategoria.SelectedItem is None and self.RowCategoria.Items.Count > 0:
            self.RowCategoria.SelectedIndex = 0

        self.cargar_familias(safe_str(self.RowCategoria.SelectedItem), self.RowFamilia, self.RowTipo)
        self.RowFamilia.SelectedItem = familia
        if self.RowFamilia.SelectedItem is None and self.RowFamilia.Items.Count > 0:
            self.RowFamilia.SelectedIndex = 0

        self.cargar_tipos(safe_str(self.RowCategoria.SelectedItem), safe_str(self.RowFamilia.SelectedItem), self.RowTipo)
        self.RowTipo.SelectedItem = tipo
        if self.RowTipo.SelectedItem is None and self.RowTipo.Items.Count > 0:
            self.RowTipo.SelectedIndex = 0

    def on_row_categoria_changed(self, sender, args):
        categoria = safe_str(self.RowCategoria.SelectedItem)
        self.cargar_familias(categoria, self.RowFamilia, self.RowTipo)

    def on_row_familia_changed(self, sender, args):
        categoria = safe_str(self.RowCategoria.SelectedItem)
        familia = safe_str(self.RowFamilia.SelectedItem)
        self.cargar_tipos(categoria, familia, self.RowTipo)

    def on_apply_type_to_row(self, sender, args):
        row = self.get_selected_data_row()
        if row is None:
            forms.alert("Selecciona una fila de la tabla.")
            return

        categoria = safe_str(self.RowCategoria.SelectedItem)
        familia = safe_str(self.RowFamilia.SelectedItem)
        tipo = safe_str(self.RowTipo.SelectedItem)

        if not categoria or not familia or not tipo:
            forms.alert("Debes seleccionar categoria, familia y tipo.")
            return

        row["categoria"] = categoria
        row["familia"] = familia
        row["tipo"] = tipo

    def recolectar_datos(self):
        self.ReviewGrid.CommitEdit()
        datos = []

        for row in self.review_table.Rows:
            fila_num = row["fila"]
            comentario = safe_str(row["comentario"]).strip()
            categoria = safe_str(row["categoria"]).strip()
            familia = safe_str(row["familia"]).strip()
            tipo = safe_str(row["tipo"]).strip()

            norte = try_parse_float(row["norte"])
            este = try_parse_float(row["este"])
            cota = try_parse_float(row["cota"])

            cond_ingreso = ""
            cond_salida = ""
            conexion = ""

            if self.usar_columnas_syp:
                cond_ingreso = safe_str(row["000_SYP_COND_INGRESO"]).strip()
                cond_salida = safe_str(row["000_SYP_COND_SALIDA"]).strip()
                conexion = safe_str(row["000_SYP_CONEXION"]).strip()

            datos.append({
                "fila": fila_num,
                "comentario": comentario,
                "norte": norte,
                "este": este,
                "cota": cota,
                "categoria": categoria,
                "familia": familia,
                "tipo": tipo,
                "000_SYP_COND_INGRESO": cond_ingreso,
                "000_SYP_COND_SALIDA": cond_salida,
                "000_SYP_CONEXION": conexion
            })

        return datos

    def validar(self, datos):
        errores = []

        for d in datos:
            if d["norte"] is None:
                errores.append("Fila {}: NORTE invalido.".format(d["fila"]))

            if d["este"] is None:
                errores.append("Fila {}: ESTE invalido.".format(d["fila"]))

            if d["cota"] is None:
                errores.append("Fila {}: COTA invalida.".format(d["fila"]))

            if not d["categoria"]:
                errores.append("Fila {}: sin categoria.".format(d["fila"]))

            if not d["familia"]:
                errores.append("Fila {}: sin familia.".format(d["fila"]))

            if not d["tipo"]:
                errores.append("Fila {}: sin tipo.".format(d["fila"]))

        return errores

    def on_review_create(self, sender, args):
        datos = self.recolectar_datos()
        errores = self.validar(datos)

        if errores:
            forms.alert("Se encontraron errores:\n\n" + "\n".join(errores[:30]))
            return

        self.resultado = datos
        self.cancelado = False
        self.Close()

    def on_cancel(self, sender, args):
        self.resultado = None
        self.cancelado = True
        self.Close()


# =========================================================
# CREACION DE ELEMENTOS
# CSV = Shared Coordinates
# =========================================================
def crear_elementos_desde_datos(doc, datos, catalogo, usar_columnas_syp=True):
    niveles = get_all_levels(doc)

    if not niveles:
        raise Exception("No se encontro ningun Level en el modelo.")

    creados = 0
    omitidos = 0
    detalle = []

    with revit.Transaction("Crear elementos desde CSV"):
        for d in datos:
            try:
                norte_m = d["norte"]
                este_m = d["este"]
                cota_m = d["cota"]

                punto = shared_a_internal_xyz(doc, norte_m, este_m, cota_m)
                nivel = obtener_nivel_mas_cercano(niveles, punto.Z)

                symbol = buscar_symbol_por_texto(
                    catalogo,
                    d["categoria"],
                    d["familia"],
                    d["tipo"]
                )

                if not symbol:
                    omitidos += 1
                    detalle.append(
                        "Fila {} [{}]: no se encontro el tipo seleccionado.".format(
                            d["fila"],
                            d["comentario"]
                        )
                    )
                    continue

                instancia = crear_instancia_generica(doc, symbol, punto, nivel)

                asignar_comentario(instancia, d["comentario"])
                asignar_nivel_si_aplica(instancia, nivel)

                if usar_columnas_syp:
                    ok_ing = asignar_parametro_texto(
                        instancia,
                        "000_SYP_COND_INGRESO",
                        d.get("000_SYP_COND_INGRESO", "")
                    )

                    ok_sal = asignar_parametro_texto(
                        instancia,
                        "000_SYP_COND_SALIDA",
                        d.get("000_SYP_COND_SALIDA", "")
                    )

                    ok_con = asignar_parametro_texto(
                        instancia,
                        "000_SYP_CONEXION",
                        d.get("000_SYP_CONEXION", "")
                    )

                    if not ok_ing:
                        detalle.append(
                            "Fila {} [{}]: no se pudo asignar 000_SYP_COND_INGRESO. Verifica que exista como parametro de instancia tipo Texto y que no sea solo lectura.".format(
                                d["fila"],
                                d["comentario"]
                            )
                        )

                    if not ok_sal:
                        detalle.append(
                            "Fila {} [{}]: no se pudo asignar 000_SYP_COND_SALIDA. Verifica que exista como parametro de instancia tipo Texto y que no sea solo lectura.".format(
                                d["fila"],
                                d["comentario"]
                            )
                        )

                    if not ok_con:
                        detalle.append(
                            "Fila {} [{}]: no se pudo asignar 000_SYP_CONEXION. Verifica que exista como parametro de instancia tipo Texto y que no sea solo lectura.".format(
                                d["fila"],
                                d["comentario"]
                            )
                        )

                creados += 1

            except Exception as ex:
                omitidos += 1
                detalle.append(
                    "Fila {} [{}]: {}".format(
                        d["fila"],
                        d["comentario"],
                        ex
                    )
                )

    return creados, omitidos, detalle


# =========================================================
# MAIN
# =========================================================
def main():
    catalogo = agrupar_simbolos(doc)

    if not catalogo:
        forms.alert(
            "No se encontraron familias cargables con tipos disponibles en el proyecto."
        )
        return

    win = ImportarCsvWindow(catalogo)
    win.ShowDialog()

    if win.cancelado or not win.resultado:
        forms.alert("Operacion cancelada.")
        return

    try:
        creados, omitidos, detalle = crear_elementos_desde_datos(
            doc,
            win.resultado,
            catalogo,
            win.usar_columnas_syp
        )
    except Exception as ex:
        forms.alert("Error general:\n{}".format(ex))
        return

    mensaje = "Elementos creados: {}\nElementos omitidos: {}".format(
        creados,
        omitidos
    )

    if detalle:
        mensaje += "\n\nDetalle:\n" + "\n".join(detalle[:30])

    forms.alert(mensaje, title="Resultado")


if __name__ == "__main__":
    main()
