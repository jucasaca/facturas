import string
import time

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
def abrir_form_gen(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	nombre = event.Source.Model.Tag
	form = XSCRIPTCONTEXT.getDocument()
	form.CurrentController.Frame.close(True)
	doc.OpenFormDocument(nombre)

	return


# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se abre un formulario
def abrir_formulario(event=None):
	ocultar_menus(event)
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
def abrir_menu_ppal(event=None):
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	doc.OpenFormDocument('MenuPpal')
	return


# ----------------------------------------------------------------------
# Actualiza el importe de las asistencias después de añadir un detalle
def actualizarImporteAsistencia(event=None):
	doc = event.Source.Parent
	doc.getByName('Totales').reload()
	return


# ----------------------------------------------------------------------
# Cargar la configuración de la tabla configuración en variables locales
def calcular_factura(event=None):
	form = event.Source.Parent  # Obtiene el formulario principal (porque hemos llamado desde SubForm)
	fa_id = form.getString(form.findColumn("FaId"))
	con = form.ActiveConnection
	stat = con.createStatement()
	sql = f'EXECUTE PROCEDURE P_CALC_FACTURA({fa_id})'
	stat.execute(sql)
	pos = form.getBookmark()
	form.reload()
	form.moveToBookmark(pos)
	return


# ----------------------------------------------------------------------
# Cargar la configuración de la tabla configuración en variables locales
def cargar_config(event=None):
	bas = CreateScriptService('Basic')
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
def cerrar_formulario(event=None):
	mostrar_menus(event)
	limpiar_filtros(event)
	bas = CreateScriptService('Basic')
	doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
	# TODO Sustituir MenuPpal por un formulario genérico
	doc.OpenFormDocument('MenuPpal')
	return


# ----------------------------------------------------------------------
# Cierra el formulario principal y muestra Base
def cerrar_menu_ppal(event=None):
	mostrar_menus(event)  # Muestra el menú y la barras nuevamente
	salir(event)
	return


# ----------------------------------------------------------------------
# Crea un registro de asistencia de colaborador con sus detalles
def crearAsistColaborador(form, as_id):
	sql = f"SELECT AC_ID FROM P_ASIST_COLABORADOR({as_id})"
	con = form.ActiveConnection
	stat = con.createStatement()
	rs = stat.executeQuery(sql)
	rs.first()
	form.reload()
	return


# ----------------------------------------------------------------------
# Ajusta el tamaño de los formularios
def establecer_tamanio(event=None):
	titulo = event.Source.Title.split(':')
	tit = titulo[1].strip()
	if tit == 'Facturas':
		w = 977
		h = 830
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
# Crea un registro de facturas de colaborador y sus detalles con los datos de la asistencia
def facturarColaborador(event=None):
	bas = CreateScriptService('Basic')
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
		mensaje('Debe seleccionar alguna fila', bas.MB_ICONINFORMATION, 'Error en la selección')
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
		imprimirFacCol(form, fact)
	form.reload()
	return


# ----------------------------------------------------------------------
# Crea un registro de factura proforma y sus detalles con los datos de la asistencia
def faturar_proforma(event=None):
	form = event.Source.Model.Parent
	fa_id = form.getInt(form.findColumn("FaId"))
	imprimirProforma(form, fa_id)
	return


# ----------------------------------------------------------------------
# Crea un registro de facturas y otro de factura de colaborador con
# sus detalles respectivos desde los datos de la asistencia
def facturar(event=None):
	bas = CreateScriptService('Basic')
	form = event.Source.Model.Parent
	if form.getBoolean(form.findColumn("FaBloqueada")):
		mensaje("No se puede facturar.\nLa asistencia ya está facturada.", bas.MB_ICONEXCLAMATION, "Error de facturación")
		return
	fa_id = form.getString(form.findColumn("FaId"))
	sql = f'EXECUTE PROCEDURE P_FACTURAR({fa_id})'
	mensaje(sql)
	con = form.ActiveConnection
	stat = con.createStatement()
	stat.executeQuery(sql)
	mensaje(fa_id)
	# imprimir_factura(form, fa_id)
	return


# ----------------------------------------------------------------------
# Crea un registro de facturas y sus detalles con los datos de la asistencia
def generar_factura(form, as_id):
	# TODO ¿ESTO HACE FALTA?
	sql = f"SELECT FA_ID FROM P_FACTURA_ASISTENCIA({as_id})"
	con = form.ActiveConnection
	stat = con.createStatement()
	rs = stat.executeQuery(sql)
	rs.first()
	fa_id = rs.getString(rs.findColumn('FA_ID'))
	imprimir_factura(form, fa_id)
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
def filtrar_no_cobradas(event=None):
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
def imprimir_factura(form, fa_id):
	bas = CreateScriptService('Basic')
	cargar_config()
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
	# informe = doc.ReportDocuments.getByName(tipoFactura).open()
	informe = bas.ThisDatabaseDocument.ReportDocuments.getByName(tipoFactura).open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)
	# Crea el path+nombre de la factura para almacenar el PDF
	archivo = uno.systemPathToFileUrl(dir_facturas + numFactura + '.pdf')
	# Imprimir la factura
	args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
	informe.storeToURL(archivo, args)
	informe.close(True)
	# Limpia el filtro para el próximo uso
	limpiar_filtros()
	return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimirFacCol(form, fc_id):
	bas = CreateScriptService('Basic')
	sql = f'UPDATE "Filtros" SET "Valor" = {fc_id} WHERE "FiId" = 1'
	con = form.ActiveConnection
	stat = con.createStatement()
	stat.executeUpdate(sql)
	# Abrir el informe y ocultarlo
	informe = bas.ThisDatabaseDocument.ReportDocuments.getByName('FacturaColaborador').open()
	# informe = doc.getByName('FacturaColaborador').open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)
	# Obtener el número de factura para ponerlo en el nombre del archivo
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
	limpiar_filtros()
	return


# ----------------------------------------------------------------------
# Reimprime una factura de colaborador desde el formulario de facturas
def imprimirFactColForm(event=None):
	form = event.Source.Model.Parent
	fc_id = form.getString(form.findColumn("FcId"))
	imprimirFacCol(form, fc_id)
	return


# ----------------------------------------------------------------------
# Reimprime una factura desde el formulario de facturas
def reimprimir_factura(event=None):
	bas = CreateScriptService('Basic')
	doc = bas.ThisDatabaseDocument
	form = event.Source.Model.Parent
	fa_id = form.getString(form.findColumn("FaId"))
	imprimir_factura(form, fa_id)
	return

# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimirProforma(form, fa_id):
	bas = CreateScriptService('Basic')
	sql = f'UPDATE "Filtros" SET "Valor" = {fa_id} WHERE "FiId" = 1'
	con = form.ActiveConnection
	stat = con.createStatement()
	stat.executeUpdate(sql)
	# abrir el informe y ocultarlo
	informe = bas.ThisDatabaseDocument.ReportDocuments.getByName('FacturaProforma').open()
	vistaInforme = informe.CurrentController.Frame.ContainerWindow
	vistaInforme.setVisible(False)
	# Nombre para el PDF creado
	archivo = uno.systemPathToFileUrl(dir_facturas + 'PROFORMA-' + fp_id + '.pdf')
	# Crear el pdf, guardarlo y cerrar el informe
	args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
	informe.storeToURL(archivo, args)
	informe.close(True)
	# Limpia el filtro para el próximo uso
	limpiar_filtros()
	return
# ----------------------------------------------------------------------
# Ejecuta las rutinas necesarias para iniciar el programa
def iniciar_programa(event=None):
	Application.OpenConnection()
	abrir_menu_ppal(event)
	ocultar_base(event)
	cargar_config(event)
	return


# ----------------------------------------------------------------------
# Pone en blanco todos los campos de la tabla Filtros (para cancelar el filtrado)
def limpiar_filtros(event=None):
	# Primero vaciar el contenido de todos los campos de la tabla auxiliar filtros
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
			# recargar todos lo formularios para que actualicen los datos y se muestren todos
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
def mostrar_base(event=None):
	DoCmd.SelectObject(acConstants.acDatabaseWindow)
	DoCmd.Maximize()
	return


# ----------------------------------------------------------------------
# Muestra el menu y barras de herramientas de un formulario que los tiene ocultos
def mostrar_menus(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	frame.LayoutManager.setVisible(True)
	pass


# ----------------------------------------------------------------------
# Oculta Base. Se llama desde un formulario
def ocultar_base(event=None):
	# DoCmd.SetHiddenAttribute(acConstants.acDatabaseWindow)
	DoCmd.SelectObject(acConstants.acDatabaseWindow)
	DoCmd.Minimize()
	return


# ----------------------------------------------------------------------
# Esconde el menu y barras de herramientas de un formulario
def ocultar_menus(event=None):
	doc = event.Source
	frame = doc.CurrentController.Frame
	frame.LayoutManager.setVisible(False)
	establecer_tamanio(event)


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
	from Xlib import display

	return

# ----------------------------------------------------------------------
def pruebas(event=None):
	form = event.Source.Model.Parent
	tabla = form.getByName('tblAsistencias')
	vista = XSCRIPTCONTEXT.getDocument().getCurrentController().getControl(tabla)
	selec = vista.getSelection()

	con = form.ActiveConnection
	stat = con.createStatement()
	# Primero vaciar la tabla de parámetros
	sql = 'DELETE FROM "Parametros" WHERE 1=1'
	stat.executeUpdate(sql)
	# Si hay selecciones guardar los id en la tabla de parámetros
	if selec:
		for s in selec:
			form.absolute(s)
			valor = form.Columns.getByName('AcId').getString()
			sql = f'INSERT INTO "Parametros" ("PaValor") VALUES ({valor})'
			stat.executeUpdate(sql)
	else:
		mensaje('Debe seleccionar alguna fila')
	return
