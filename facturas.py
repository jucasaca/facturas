import string
from com.sun.star.document import DocumentEvent
from com.sun.star.beans import PropertyValue
from scriptforge import CreateScriptService
from access2base import DoCmd, Application, acConstants, THISDATABASEDOCUMENT
import uno

dir_facturas = ''
dir_fac_colab = ''


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
# Rutinas a ejecutar cuando se abre un formulario
def abrirFormulario(event=None):
	ocultarMenuBarras(event)
	return


# ----------------------------------------------------------------------
# Abre un informe. El nombre del informe debe estar en el Tag del 
# control que lo llama
def abrirInformeGenerico(event=None):
	nombre = event.Source.Model.Tag
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	repo = doc.ReportDocuments.getByName(nombre)
	# doc.ReportDocuments.getByName(nombre).open()
	return


# ----------------------------------------------------------------------
# Muestra el formulario principal y oculta Base
def abrirMenuPpal(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	doc.OpenFormDocument('MenuPpal')
	return



# ----------------------------------------------------------------------
# Cargar la configuración de la tabla configuración en variables locales
def cargarConfig(event=None):
	bas = CreateScriptService(('Basic'))
	ds = bas.thisDatabaseDocument.DataSource
	con = ds.getConnection('','')
	stat = con.createStatement()

	# Obtener el directorio de facturas
	sql = """SELECT "CfValor" FROM "Configuracion" WHERE "CfConfiguracion" = 'DirFact'"""
	rs = stat.executeQuery(sql)
	rs.first()
	global dir_facturas
	dir_facturas = rs.getString(rs.findColumn('CfValor'))

	# Obtener el directorio de facturas de colaborador
	sql = """SELECT "CfValor" FROM "Configuracion" WHERE "CfConfiguracion" = 'DirFactCol'"""
	rs = stat.executeQuery(sql)
	rs.first()
	global dir_fac_colab
	dir_fac_colab = rs.getString(rs.findColumn('CfValor'))

	return

# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se cierra un formulario
def cerrarFormulario(event=None):
	mostrarMenuBarras(event)
	# TODO ver si se puede evitar llamar siempre a limpiarFiltros
	limpiarFiltros(event)
	return


# ----------------------------------------------------------------------
# Cierra el formulario principal y muestra Base
def cerrarMenuPpal(event=None):
	mostrarMenuBarras(event)  # Muestra el menú y la barras nuevamente
	salir(event)
	return

# ----------------------------------------------------------------------
# Ajusta el tamaño de los formularios
def establecerTamanio(event=None):
	titulo = event.Source.Title.split(':')
	tit = titulo[1].strip()
	if tit == 'Facturas':
		w = 960
		h = 690
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
	elif tit == 'Asistencias':
		w = 970
		h = 750
	elif tit == 'Colaboradores':
		w = 665
		h = 560
	elif tit == 'FacturasColaborador':
		w = 960
		h = 730
	elif tit == 'Configuracion':
		w = 740
		h = 480
	else:
		w = -1
		h = -1
		mensaje('Tamaño->No se ha encontrado el formulario')
	Application.OpenConnection()
	DoCmd.MoveSize(width=w, height=h)


# ----------------------------------------------------------------------
# Emite la factura (pone número e imprime)
def emitirFactura(fa_id):
	bas = CreateScriptService(('Basic'))
	ds = bas.thisDatabaseDocument.DataSource
	con = ds.getConnection('', '')
	stat = con.createStatement()
	sSQL = 'EXECUTE PROCEDURE P_EMITIR_FACTURA(' + fa_id + ')'
	stat.executeQuery(sSQL)

	imprimirFactura(fa_id)
	return

# ----------------------------------------------------------------------
# Crea un registro de facturas y sus detalles con los datos de la asistencia
def facturarAsistencia(event=None):
	form = event.Source.Model.Parent
	if form.getInt(form.findColumn("AsIdFactura")):
		mensaje("No se puede facturar la asistencia.\nLa asistencia ya está facturada.", 48, "Error de facturación")
		return
	id = form.getString(form.findColumn("AsId"))
	sSQL = "SELECT FA_ID FROM P_FACTURA_ASISTENCIA(" + id + ")"
	con = form.ActiveConnection
	stat = con.createStatement()
	rs = stat.executeQuery(sSQL)
	rs.first()
	fa_id = rs.getString(rs.findColumn('FA_ID'))
	registro_actual = form.getBookmark()
	form.reload()
	form.moveToBookmark(registro_actual)

	emitirFactura(fa_id)
	return


# ----------------------------------------------------------------------
# Crea un registro de factura de colaborador con sus detalles
def facturarColaborador(event=None):
	form = event.Source.Model.Parent
	if form.getInt(form.findColumn("AsIdFactColaborador")):
		mensaje("""No se puede hacer factura de colaborador.\n\
		La asistencia ya está facturada a un colaborador.""", 48, "Error de facturación")
		return
	id = form.getString(form.findColumn("AsId"))
	sSQL = "EXECUTE PROCEDURE P_FACT_COLABORADOR(" + id + ")"
	con = form.ActiveConnection
	stat = con.prepareStatement(sSQL)
	stat.executeQuery()
	registro_actual = form.getBookmark()
	form.reload()
	form.moveToBookmark(registro_actual)
	return


# ----------------------------------------------------------------------
# Crea un registro de facturas y otro de factura de colaborador con
# sus detalles respectivos desde los datos de la asistencia
def facturarTodo(event=None):
	facturarAsistencia(event)
	facturarColaborador(event)
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no emitidas en el formulario facturas
def filtrarAsistencias(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		boton.HelpText = "Mostrar todas las asistencias"
		form.Filter = """"AsIdFactura" IS NULL OR "AsIdFactColaborador" IS  NULL"""
		form.ApplyFilter = True
		form.reload()
	else:
		boton.HelpText = "Mostrar solo asistencias no facturadas"
		form.Filter = ""
		form.reload()
	return

# ----------------------------------------------------------------------
# Establece un filtro de facturas no emitidas en el formulario facturas
def filtrarColabNoFacturadas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		if form.getByName("btnNoPagadas").State:
			form.getByName("btnNoPagadas").State = 0
		form.Filter = "FcNumero IS NULL"
		form.ApplyFilter = True
		form.reload()
	else:
		form.Filter = ""
		form.reload()
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no pagadas en el formulario facturas de colaborador
def filtrarColabNoPagadas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		if form.getByName("btnNoFacturadas").State:
			form.getByName("btnNoFacturadas").State = 0
		form.Filter = "FcFechaPago IS NULL"
		form.ApplyFilter = True
		form.reload()
	else:
		form.Filter = ""
		form.reload()
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no cobradas en el formulario facturas
def filtrarNoCobradas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		if form.getByName("btnNoFacturadas").State:
			form.getByName("btnNoFacturadas").State = 0
		form.Filter = "FaFechaCobro IS NULL"
		form.ApplyFilter = True
		form.reload()
	else:
		form.Filter = ""
		form.reload()
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no emitidas en el formulario facturas
def filtrarNoFacturadas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		if form.getByName("btnNoCobradas").State:
			form.getByName("btnNoCobradas").State = 0
		form.Filter = "FaNumero IS NULL"
		form.ApplyFilter = True
		form.reload()
	else:
		form.Filter = ""
		form.reload()
	return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimirFactura(fa_id):
	bas = CreateScriptService('Basic')
	# Filtrar elinforme por el Id de factura
	sql = 'UPDATE "Filtros" SET "Valor" = ' + fa_id + ' WHERE "FiId" = 1'
	ds = bas.thisDatabaseDocument.DataSource
	con = ds.getConnection('', '')
	stat = con.createStatement()
	stat.executeUpdate(sql)
	# Abrir el informe y ocultarlo
	doc = bas.ThisDatabaseDocument
	informe = doc.ReportDocuments.getByName("FacturaGeneral").open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)
	# Obtener el numero de factura para ponerlo en el nombre del archivo
	sql = 'SELECT "FaNumero" FROM "Facturas" WHERE "FaId" = ' + fa_id
	rs = stat.executeQuery(sql)
	rs.first()
	numFactura = rs.getString(rs.findColumn('FaNumero'))
	archivo = uno.systemPathToFileUrl(dir_facturas + numFactura + '.pdf')
	# Imprimir la factura
	args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
	informe.storeToURL(archivo, args)
	informe.close(True)
	# Limpia el filtro para el próximo uso
	limpiarFiltros()
	return


# ----------------------------------------------------------------------
# Ejecuta las rutinas necesariae para iniciar el programa
def imprimirFacturaForm(event=None):
	# TODO obtener el id de la factura seleccionada e imprimirla
	mensaje("Hay que ver somo se llama a imprimirFacturad desde aquí", 48, "Error")
	return


# ----------------------------------------------------------------------
# Ejecuta las rutinas necesariae para iniciar el programa
def iniciarPrograma(event=None):
	Application.OpenConnection()
	cargarConfig(event)
	abrirMenuPpal(event)
	# # TODO: Ocultar Base según necesidades
	ocultarBase(event)
	return


# ----------------------------------------------------------------------
# Pone en blanco todos los campos de la tabla Filtros (para cancelar el filtrado)
def limpiarFiltros(event=None):
	# Primero vacía el contenido de todos los campos de la tabal auxiliar filtros
	# Application.OpenConnection()
	rs = Application.CurrentDb().OpenRecordset("Filtros")
	rs.Edit()
	for f in rs.Fields():
		if f.Name != 'FiId':
			f.Value = ''
	rs.Update()
	if event:
		source = event.Source  # ¿Quién llama a la función?
		# Si es un botón
		if source.ImplementationName == 'com.sun.star.form.OButtonControl':
			# recargamos todos lo formularios para que actualicen los datos y se muestren todos
			for form in source.Model.Parent.Parent:  # la colección de formularios
				form.reload()
	return


# ----------------------------------------------------------------------
# Ventana de mensajes tipo MsgBox
def mensaje(texto, botones=0, titulo=''):
	bas = CreateScriptService("Basic")
	return bas.MsgBox(texto, botones, titulo)
# ----------------------------------------------------------------------
# Muestra Base.


def mostrarBase(event=None):
	# DoCmd.SetHiddenAttribute(acConstants.acDatabaseWindow, hidden=False)
	DoCmd.SelectObject(acConstants.acDatabaseWindow)
	DoCmd.Maximize()
	return


# ----------------------------------------------------------------------
# Muestra el menu y barras de herramientas de un formulario que los tiene ocultos
def mostrarMenuBarras(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	frame.LayoutManager.setVisible(True)
	pass


# ----------------------------------------------------------------------
# Oculta Base. Se llama desde un formulario
def ocultarBase(event=None):
	# DoCmd.SetHiddenAttribute(acConstants.acDatabaseWindow)
	DoCmd.SelectObject(acConstants.acDatabaseWindow)
	DoCmd.Minimize()
	return


# ----------------------------------------------------------------------
# Esconde el menu y barras de herramientas de un formulario
def ocultarMenuBarras(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	#  TODO Cambiar la visibilidad para editar el formulario
	frame.LayoutManager.setVisible(False)
	establecerTamanio(event)


# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se cierra el programa
def salir(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService("SFDocuments.Document", bas.ThisDatabaseDocument)
	# TODO Mostar base al cerrar la aplicación
	mostrarBase(event)
	# TODO Comentar temporalmente la siguiente línea si se necesita trabajar en Base
	# doc.RunCommand("CloseDoc")
	return


# ----------------------------------------------------------------------
# Ventana de mensajes tipo MsgBox
def xray(objeto):
	bas = CreateScriptService("Basic")
	bas.Xray(objeto)


# ----------------------------------------------------------------------
def main(event=None):
	titulo = event.Source.Title.split()
	titulo = titulo.split(':')
	return

# ----------------------------------------------------------------------
def pruebas(event=None):
	mensaje('Estamos saliendo')

	return
