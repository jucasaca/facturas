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
	form = XSCRIPTCONTEXT.getDocument()
	form.CurrentController.Frame.close(True)
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


def actualizarImporteAsistencia(event=None):
	bas = CreateScriptService('Basic')
	doc = event.Source.Parent
	doc.getByName('Totales').reload()
	# xray(formImporte)
	# formImporte.reload()
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
	limpiarFiltros(event)
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	# TODO Sustituir MenuPpal por un formulario genérico
	doc.OpenFormDocument('MenuPpal')
	return


# ----------------------------------------------------------------------
# Cierra el formulario principal y muestra Base
def cerrarMenuPpal(event=None):
	mostrarMenuBarras(event)  # Muestra el menú y la barras nuevamente
	salir(event)
	return


# ----------------------------------------------------------------------
# Crea un registro de asistencia de colaborador con sus detalles
def crearAsistColaborador(doc, form, as_id):
	sql = f"SELECT AC_ID FROM P_ASIST_COLABORADOR({as_id})"
	con = form.ActiveConnection
	stat = con.createStatement()
	rs = stat.executeQuery(sql)
	rs.first()
	# fc_id = rs.getString(rs.findColumn('AC_ID'))
	# 	imprimirFacCol(doc, form, fc_id)
	form.reload()
	return


# ----------------------------------------------------------------------
# Ajusta el tamaño de los formularios
def establecerTamanio(event=None):
	titulo = event.Source.Title.split(':')
	tit = titulo[1].strip()
	if tit == 'Facturas':
		w = 960
		h = 760
	elif tit == 'Clientes':
		w = 748
		h = 690
	elif tit == 'MenuPpal':
		w = 545
		h = 700
	elif tit == 'Gastos':
		w =1080
		h = 685
	elif tit == 'Proveedores':
		w = 748
		h = 685
	elif tit == 'SeriesFactura':
		w = 638
		h = 450
	elif tit == 'Asistencias':
		w = 973
		h = 790
	elif tit == 'Colaboradores':
		w = 685
		h = 485
	elif tit == 'AsistenciasColaborador':
		w = 880
		h = 720
	elif tit == 'FacturasColaboradores':
		w = 880
		h = 720
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
# Crea un registro de facturas y sus detalles con los datos de la asistencia
def facturarAsistencia(doc, form, as_id):
	sql = f"SELECT FA_ID FROM P_FACTURA_ASISTENCIA({as_id})"
	con = form.ActiveConnection
	stat = con.createStatement()
	rs = stat.executeQuery(sql)
	rs.first()
	fa_id = rs.getString(rs.findColumn('FA_ID'))
	imprimirFactura(doc, form, fa_id)
	return


# ----------------------------------------------------------------------
# Crea un registro de facturas de colaborador y sus detalles con los datos de la asistencia
def facturarColaborador(event=None):
	form = event.Source.Model.Parent
	doc = XSCRIPTCONTEXT.getDocument()
	tabla = form.getByName('tblAsistencias')
	vista = doc.getCurrentController().getControl(tabla)
	selec = vista.getSelection()

	con = form.ActiveConnection
	stat = con.createStatement()
	sql = 'DELETE FROM "Parametros" WHERE 1=1'
	stat.executeUpdate(sql)
	if not selec:  # Si no hay selección, sale de la función
		mensaje('Debe seleccionar alguna fila')
		return
	for s in selec:
		form.absolute(s)
		valor = form.Columns.getByName('AcId').getString()
		sql = f'INSERT INTO "Parametros" ("PaValor") VALUES ({valor})'
		stat.executeUpdate(sql)
	sql = 'SELECT FC_ID FROM P_FACT_COLAB'
	rs = stat.executeQuery(sql)
	while rs.next():
		fact = rs.getString(rs.findColumn('FC_ID'))
		imprimirFacCol(doc, form, fact)
	form.reload()
	return

# ----------------------------------------------------------------------
# Crea un registro de facturas y otro de factura de colaborador con
# sus detalles respectivos desde los datos de la asistencia
def facturarTodo(event=None):
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	form = event.Source.Model.Parent
	if form.getInt(form.findColumn("AsIdFactura")):
		mensaje("No se puede facturar la asistencia.\nLa asistencia ya está facturada.", 48, "Error de facturación")
		return
	as_id = form.getString(form.findColumn("AsId"))
	facturarAsistencia(doc, form, as_id)
	crearAsistColaborador(doc, form, as_id)
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no emitidas en el formulario facturas
def filtrarAsistencias(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		boton.HelpText = "Mostrar todas las asistencias"
		form.ApplyFilter = True
		form.reload()
	else:
		boton.HelpText = "Mostrar solo asistencias no facturadas"
		form.ApplyFilter = False
		form.reload()
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no emitidas en el formulario facturas
def filtrarAsistenciasColab(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		boton.HelpText = "Mostrar todas las asistencias"
		form.ApplyFilter = True
		form.reload()
		pass
	else:
		boton.HelpText = "Mostrar solo asistencias no facturadas"
		form.ApplyFilter = False
		form.reload()
		pass
	# return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no pagadas en el formulario facturas de colaborador
def filtrarColabNoPagadas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		boton.HelpText = "Mostrar todas las facturas"
		form.ApplyFilter = True
		form.reload()
	else:
		boton.HelpText = "Mostrar solo no pagadas"
		form.ApplyFilter = False
		form.reload()
	return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no cobradas en el formulario facturas
def filtrarNoCobradas(event=None):
	boton = event.Source.Model
	form = event.Source.Model.Parent
	if boton.State:
		boton.HelpText = "Mostrar todas las facturas"
		form.ApplyFilter = True
		form.reload()
	else:
		boton.HelpText = "Mostrar solo no cobradas"
		form.ApplyFilter = False
		form.reload()
	return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimirFactura(doc, form, fa_id):
	bas = CreateScriptService('Basic')
	cargarConfig()
	# Filtrar el informe por el Id de factura
	sql = 'UPDATE "Filtros" SET "Valor" = ' + fa_id + ' WHERE "FiId" = 1'
	con = form.ActiveConnection
	stat = con.createStatement()
	stat.executeUpdate(sql)
	# Obtener el número de factura y el concepto
	sql = f'SELECT "FaNumero", "FaConcepto" FROM "Facturas" WHERE "FaId" = {fa_id}'
	rs = stat.executeQuery(sql)
	rs.first()
	numFactura = rs.getString(rs.findColumn('FaNumero'))
	concepto = rs.getString(rs.findColumn('FaConcepto'))
	# Si tiene un concepto, se imprime la factura de concepto, si no la general
	if concepto:
		tipoFactura = 'FacturaConcepto'
	else:
		tipoFactura = 'FacturaGeneral'
	# Abrir el informe y ocultarlo
	informe = doc.ReportDocuments.getByName(tipoFactura).open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)


	archivo = uno.systemPathToFileUrl(dir_facturas + numFactura + '.pdf')
	# Imprimir la factura
	args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
	informe.storeToURL(archivo, args)
	informe.close(True)
	# Limpia el filtro para el próximo uso
	limpiarFiltros()
	return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimirFacCol(doc, form, fc_id):
	bas = CreateScriptService('Basic')
	sql = f'UPDATE "Filtros" SET "Valor" = {fc_id} WHERE "FiId" = 1'
	con = form.ActiveConnection
	stat = con.createStatement()
	stat.executeUpdate(sql)
	informe = bas.ThisDatabaseDocument.ReportDocuments.getByName('FacturaColaborador').open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)
	# Obtener el numero de factura para ponerlo en el nombre del archivo
	sql = f'SELECT "FcNumero" FROM "FacturasColaboradores" WHERE "FcId" = {fc_id}'
	rs = stat.executeQuery(sql)
	rs.first()
	numFactura = rs.getString(rs.findColumn('FcNumero'))
	archivo = uno.systemPathToFileUrl(dir_fac_colab + numFactura + '.pdf')
	# Imprimir la factura
	args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
	informe.storeToURL(archivo, args)
	informe.close(True)
	# Limpia el filtro para el próximo uso
	limpiarFiltros()
	return


# ----------------------------------------------------------------------
# Reimprime una factura de colaborador desde el formulario de facturas
def imprimirFactColForm(event=None):
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	form = event.Source.Model.Parent
	fc_id = form.getString(form.findColumn("FcId"))
	imprimirFacCol(doc, form, fc_id)
	return


# ----------------------------------------------------------------------
# Reimprime una factura desde el formulario de facturas
def imprimirFacturaForm(event=None):
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	form = event.Source.Model.Parent
	fa_id = form.getString(form.findColumn("FaId"))
	imprimirFactura(doc, form, fa_id)
	return


# ----------------------------------------------------------------------
# Ejecuta las rutinas necesariae para iniciar el programa
def iniciarPrograma(event=None):
	Application.OpenConnection()
	abrirMenuPpal(event)
	ocultarBase(event)
	cargarConfig(event)
	return


# ----------------------------------------------------------------------
# Pone en blanco todos los campos de la tabla Filtros (para cancelar el filtrado)
def limpiarFiltros(event=None):
	# Primero vacía el contenido de todos los campos de la tabla auxiliar filtros
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
	frame.LayoutManager.setVisible(False)
	establecerTamanio(event)


# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se cierra el programa
def salir(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService("SFDocuments.Document", bas.ThisDatabaseDocument)
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
	form = event.Source.Model.Parent
	tabla = form.getByName('tblAsistencias')
	vista = XSCRIPTCONTEXT.getDocument().getCurrentController().getControl(tabla)
	selec = vista.getSelection()

	con = form.ActiveConnection
	stat = con.createStatement()
	sql = 'DELETE FROM "Parametros" WHERE 1=1'
	stat.executeUpdate(sql)
	if selec:
		cad = ''
		for s in selec:
			form.absolute(s)
			valor = form.Columns.getByName('AcId').getString()
			sql = f'INSERT INTO "Parametros" ("PaValor") VALUES ({valor})'
			stat.executeUpdate(sql)
	else:
		mensaje('Debe seleccionar alguna fila')

	return
