import string
from com.sun.star.document import DocumentEvent
from scriptforge import CreateScriptService
from access2base import DoCmd, Application, acConstants
import uno


# ----------------------------------------------------------------------
# Muestra el formulario principal y oculta Base
def abrirMenuPpal(event=None):
	Application.OpenConnection()
	# # TODO: Ocultar Base
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document',bas.ThisDatabaseDocument)
	doc.OpenFormDocument('MenuPpal')
	return


# ----------------------------------------------------------------------
# Cierra el formulario principal y muestra Base
def cerrarMenuPpal(event=None):
	# TODO Mostar base al cerrar la aplicación
	mostrarMenuBarras(event)  # Muestra el menú y la barras nuevamente
	salir(event)
	return


# ----------------------------------------------------------------------
# Abre un formulario. El nombre del formulario debe estar en el Tag del 
# control que lo llama
def abrirFormGenerico(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	nombre = event.Source.Model.Tag
	doc.OpenFormDocument(nombre)
	return


# ----------------------------------------------------------------------
# Abre un informe. El nombre del informe debe estar en el Tag del 
# control que lo llama
def abrirInformeGenerico(event=None):
	nombre = event.Source.Model.Tag
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	doc.ReportDocuments.getByName(nombre).open()
	return


# ----------------------------------------------------------------------
# Oculta Base. Se llama desde un formulario
def ocultarBase(event=None):
	Application.OpenConnection()
	DoCmd.SetHiddenAttribute(acConstants.acDatabaseWindow)
	return


# ----------------------------------------------------------------------
# Muestra Base.
def mostrarBase(event=None):
	Application.OpenConnection()
	DoCmd.SetHiddenAttribute(acConstants.acDatabaseWindow, hidden=False)
	return


def abrirFormulario(event=None):
	# ocultarMenuBarras(event)
	pass

def cerrarFormulario(event=None):
	mostrarMenuBarras(event)
	# TODO ver si se puede evitar llamar siempre a limpiarFiltros
	limpiarFiltros(event)


def limpiarFiltros(event=None):
	# Primero vacía el contenido de todos los campos de la tabal auxiliar filtros
	rs = Application.CurrentDb().OpenRecordset("Filtros")
	rs.Edit()
	for f in rs.Fields():
		if f.Name != 'FiId':
			f.Value = ""
	rs.Update()
	source = event.Source  # ¿Quién llama a la función?
	# Si es un botón
	if source.ImplementationName == 'com.sun.star.form.OButtonControl':
		# recargamos todos lo formularios para que actualicen los datos y se muestren todos
		for form in source.Model.Parent.Parent:  # la colección de formularios
			form.reload()
	return


# ----------------------------------------------------------------------
# Esconde el menu y barras de herramientas de un formulario
def ocultarMenuBarras(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	#  TODO Cambiar la visibilidad para editar el formulario
	frame.LayoutManager.setVisible(False)
	tamanio(event)


# ----------------------------------------------------------------------
# Muestra el menu y barras de herramientas de un formulario que los tiene ocultos
def mostrarMenuBarras(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	frame.LayoutManager.setVisible(True)
	pass


# ----------------------------------------------------------------------
def salir(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService("SFDocuments.Document", bas.ThisDatabaseDocument)
	# TODO Comentar temporalmente la siguiente línea si se necesita trabajar en Base
	# doc.RunCommand("CloseDoc")

# ----------------------------------------------------------------------
# Ajusta el tamaño de los formularios
def tamanio(event=None):
	titulo = event.Source.Title.split(':')
	tit = titulo[1].strip()
	if tit == 'Facturas':
		w = 885
		h = 720
	elif tit == 'Clientes':
		w = 665
		h = 530
	elif tit == 'MenuPpal':
		w = 545
		h = 700
	elif tit == 'Gastos':
		w =1080
		h = 685
	elif tit == 'Proveedores':
		w = 645
		h = 505
	elif tit == 'SeriesFactura':
		w = 638
		h = 450
	elif tit == 'Asistencias' or tit == 'Asistencias1':
		w = 970
		h = 750
	else:
		w = -1
		h = -1
		mensaje('Tamaño->No se ha encontrado el formulario')
	Application.OpenConnection()
	DoCmd.MoveSize(width=w, height=h)


# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
def main(event=None):
	titulo = event.Source.Title.split()
	titulo = titulo.split(':')
	return


def facturaDesdeAsistencia(event=None):
	form = event.Source.Model.Parent
	if not form.getInt(form.findColumn("AsIdFactura")):
		id = form.getString(form.findColumn("AsId"))
		sSQL = "SELECT * FROM  P_FACTURA_DESDE_ASITENCIA(" + id + ")"
		con = form.ActiveConnection
		stat = con.prepareStatement(sSQL)
		result = stat.executeQuery()
		result.next() # al moverse al siguiente registro guarda los cambios efectuados
		registro_actual = form.getBookmark()
		form.reload()
		form.moveToBookmark(registro_actual)
	else:
		mensaje("No se puede facturar la asistencia.\nLa asistencia ya está facturada.",48,"Error de facturación")
	# mensaje(fact)
	return

def pruebas(event=None):
	bas = CreateScriptService("Basic")
	form = event.Source.Model.Parent

	# campo = form.Columns.getByName("AsId").getString()
	id = form.getString(form.findColumn("AsId"))
	# campo = form.getString(1)

	sSQL = "SELECT * FROM  P_FACTURA_DESDE_ASITENCIA(" + id + ")"
	mensaje(sSQL)

	con = form.ActiveConnection
	stat = con.prepareStatement(sSQL)
	result = stat.executeQuery()
	result.next()
	mensaje(result.getString(1))

	sSQL = """SELECT * FROM "DetallesAsistencia" WHERE 	"DaIdAsistencia" = """ + campo

	con = form.ActiveConnection
	stat = con.prepareStatement()

	ret = stat.executeQuery(sSQL)
	# mensaje(ret)
	rs = stat.executeQuery(sSQL)
	rs.first()
	while not rs.isAfterLast():
		mensaje(rs.getString(rs.findColumn("DaDescripcion")))
		rs.next()
		# sSQL = """ INSERT INTO """
	return

def obtenerCampo(event, nombreCampo):
	form = event.Source.Model.Parent
	campo = form.getString(form.findColumn(nombreCampo))
	return campo
def mensaje(texto, botones=0, titulo=''):
	bas = CreateScriptService("Basic")
	bas.MsgBox(texto, botones, titulo)


def xray(objeto):
	bas = CreateScriptService("Basic")
	bas.Xray(objeto)
