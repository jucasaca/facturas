import string
import time

from com.sun.star.document import DocumentEvent
from com.sun.star.beans import PropertyValue
from scriptforge import CreateScriptService
import uno

dir_facturas = ''
dir_fac_colab = ''


# ----------------------------------------------------------------------
# Abre un formulario. El nombre del formulario debe estar en el Tag del
# control que lo llama
def abrir_form_gen(event=None):
    bas = CreateScriptService('Basic')
    doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
    # Obtiene el nombre del formulario a abrir de la etiqueta del ¿botón? que lo llama
    nombre = event.Source.Model.Tag
    # Obtiene el nombre del formulario actual a partir del título de la ventana
    ui = CreateScriptService('UI')
    ventana = ui.ActiveWindow.split(':')
    titulo = ventana[1].strip()
    # Abre el nuevo formulario
    doc.OpenFormDocument(nombre)
    # Cierra el formulario actual
    bas.ThisDatabaseDocument.FormDocuments.getByName(titulo).close()
    return

# ----------------------------------------------------------------------
# Abre un INFORME. El nombre del INFORME debe estar en el Tag del
# control que lo llama
def abrir_report_gen(event=None):
    bas = CreateScriptService('Basic')
    doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
    # Obtiene el nombre del informe a abrir de la etiqueta del ¿botón? que lo llama
    nombre = event.Source.Model.Tag
    
    # Abrir el informe y ocultarlo
    informe = bas.ThisDatabaseDocument.ReportDocuments.getByName(nombre).open()
    # vistaInforme = informe.CurrentController.Frame.ContainerWindow
    # vistaInforme.setVisible(False)
    
    return


# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se abre un formulario
def abrir_formulario(event=None):
    ocultar_menus(event)
    establecer_tamanio(event)
    limpiar_filtros(event)
    return


# ----------------------------------------------------------------------
# Muestra el formulario principal y oculta Base
def abrir_menu_ppal(event=None):
    bas = CreateScriptService('Basic')
    doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
    doc.OpenFormDocument('MenuPpal')
    return


# ----------------------------------------------------------------------
# Cargar la configuración de la tabla configuración en variables locales
def refrescar_factura(event=None):
    form = event.Source.Parent  # Obtiene el formulario principal (porque hemos llamado desde SubForm)
    pos = form.getBookmark()
    form.reload()
    form.moveToBookmark(pos)
    return


# ----------------------------------------------------------------------
# Cargar la configuración de la tabla configuración en variables locales
def cargar_config(event=None):
    bas = CreateScriptService('Basic')
    ds = bas.thisDatabaseDocument.DataSource
    con = ds.getConnection('', '')
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
    bas = CreateScriptService('Basic')
    doc = CreateScriptService('Document', bas.ThisDatabaseDocument)
    mostrar_menus(event)
    # TODO Sustituir MenuPpal por un formulario genérico
    doc.OpenFormDocument('MenuPpal')
    return


# ----------------------------------------------------------------------
# Cierra el formulario principal y muestra Base
def cerrar_menu_ppal(event=None):
    mostrar_menus(event)  # Muestra el menú y la barras nuevamente
    # salir(event)
    return


# ----------------------------------------------------------------------
# Ajusta el tamaño de los formularios
def establecer_tamanio(event=None):
    ui = CreateScriptService('UI')
    ventana = ui.ActiveWindow.split(':')
    titulo = ventana[1].strip()
    x = -1
    y = -1
    w = -1
    h = -1
    
    if titulo == 'Facturas':
        w = 938
        h = 690
        x = 220
        y = 35
    elif titulo == 'Clientes':
        w = 710
        h = 620
        x = 350
        y = 70
    elif titulo == 'MenuPpal':
        w = 425
        h = 570
        x = 450
        y = 80
    elif titulo == 'Gastos':
        w = 1035
        h = 648
        x = 170
        y = 50
    elif titulo == 'Proveedores':
        w = 710
        h = 585
        x = 350
        y = 80
    elif titulo == 'SeriesFactura':
        w = 638
        h = 445
        x = 350
        y = 80
    elif titulo == 'Colaboradores':
        w = 650
        h = 470
        x = 350
        y = 80
    elif titulo == 'AstColab':
        w = 920
        h = 597
        x = 230
        y = 70
    elif titulo == 'FacturasColaborador':
        w = 837
        h = 667
        x = 260
        y = 40
    elif titulo == 'Configuracion':
        w = 730
        h = 475
        x = 350
        y = 80
    # ~ else:
        # ~ mensaje(f'Tamaño->No se ha encontrado el formulario {titulo}')

    ui.Resize(width=w, height=h, left= x, top=y)

# ----------------------------------------------------------------------
# Crea un registro de facturas de colaborador y sus detalles con los datos de la asistencia
def facturar_colaborador(event=None):
    bas = CreateScriptService('Basic')
    form = event.Source.Model.Parent
    doc = XSCRIPTCONTEXT.getDocument()
    tabla = form.getByName('tblAst')
    vista = doc.getCurrentController().getControl(tabla)
    selec = vista.getSelection()

    con = form.ActiveConnection
    stat = con.createStatement()
    sql = 'DELETE FROM "Parametros" WHERE 1=1'
    stat.executeUpdate(sql)
    if not selec:  # Si no hay selección, sale de la función
        bas.MsgBox('Debe seleccionar alguna fila', bas.MB_ICONINFORMATION, 'Error en la selección')
        return
    for s in selec:
        form.absolute(s)
        valor = form.Columns.getByName('DfId').getString()
        sql = f'INSERT INTO "Parametros" ("PaValor") VALUES ({valor})'
        stat.executeUpdate(sql)
    sql = 'SELECT FC_ID FROM P_FACT_COLAB'
    rs = stat.executeQuery(sql)
    while rs.next():
        fact = rs.getString(rs.findColumn('FC_ID'))
        imprimir_colaborador(form, fact)
        bas.MsgBox(f'Se ha imprimido la factura\ncon identificador {fact}', bas.MB_ICONINFORMATION, 'Facturar')
    form.reload()
    return


# ----------------------------------------------------------------------
# Crea un registro de factura proforma y sus detalles con los datos de la asistencia
def faturar_proforma(event=None):
    form = event.Source.Model.Parent
    fa_id = form.getInt(form.findColumn("FaId"))
    imprimir_proforma(form, fa_id)
    return


# ----------------------------------------------------------------------
# Crea un registro de facturas y otro de factura de colaborador con
# sus detalles respectivos desde los datos de la asistencia
def facturar(event=None):
    bas = CreateScriptService('Basic')
    form = event.Source.Model.Parent
    if form.getBoolean(form.findColumn("FaBloqueada")):
        mensaje("No se puede facturar.\nLa asistencia ya está facturada.", bas.MB_ICONEXCLAMATION,
                "Error de facturación")
        return
    fa_id = form.getString(form.findColumn("FaId"))
    sql = f'EXECUTE PROCEDURE P_FACTURAR({fa_id})'
    con = form.ActiveConnection
    stat = con.createStatement()
    stat.executeQuery(sql)
    pos = form.getBookmark()
    form.reload()
    form.moveToBookmark(pos)
    imprimir_factura(form, fa_id)
    return


# ----------------------------------------------------------------------
# Establece un filtro de facturas no pagadas en el formulario facturas de colaborador
def filtrar_colab_no_pagadas(event=None):
    boton = event.Source.Model
    form = event.Source.Model.Parent
    if boton.State:
        boton.HelpText = "Mostrar todas las facturas"
        form.Filter = '"FcFPago" IS NULL'
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
        boton.HelpText = 'Mostrar todas las facturas'
        form.Filter = '"FaFCobro" IS NULL'
        form.ApplyFilter = True
        form.reload()
    else:
        boton.HelpText = 'Mostrar solo no cobradas'
        form.ApplyFilter = False
        form.reload()
    return

def imprimir_fact_col(event=None):
    form = event.Source.Model.Parent
    fact = form.getString(form.findColumn('FcId'))
    imprimir_colaborador(form, fact)
    return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimir_factura(form, fa_id):
    bas = CreateScriptService('Basic')
    cargar_config()

    # Filtrar para el informe por el Id de factura
    sql = 'UPDATE "Filtros" SET "Valor" = ' + fa_id + ' WHERE "FiId" = 1'
    con = form.ActiveConnection
    stat = con.createStatement()
    stat.executeUpdate(sql)

    # Obtener el número de factura y el concepto
    numFactura = form.getString(form.findColumn('FaNumero'))
    concepto = form.getString(form.findColumn('FaConcepto'))
    # Si tiene un concepto, se imprime la factura de concepto, si no la general
    if concepto:
        tipo_factura = 'FacturaConcepto'
    else:
        tipo_factura = 'FacturaGeneral'
    # Abrir el informe y ocultarlo
    informe = bas.ThisDatabaseDocument.ReportDocuments.getByName(tipo_factura).open()
    vistaInforme = informe.CurrentController.Frame.ContainerWindow
    vistaInforme.setVisible(False)
    # Crea el path+nombre de la factura para almacenar el PDF
    archivo = uno.systemPathToFileUrl(dir_facturas + numFactura + '.pdf')
    # Imprimir la factura
    args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
    informe.storeToURL(archivo, args)
    informe.close(True)
    # Limpia el filtro para el próximo uso
    #limpiar_filtros()
    return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimir_colaborador(form, fc_id, doc=None):
    bas = CreateScriptService('Basic')
    sql = f'UPDATE "Filtros" SET "Valor" = {fc_id} WHERE "FiId" = 1'
    con = form.ActiveConnection
    stat = con.createStatement()
    stat.executeUpdate(sql)
    # Abrir el informe y ocultarlo
    informe = bas.ThisDatabaseDocument.ReportDocuments.getByName('FacturaColaborador').open()

    vistaInforme = informe.CurrentController.Frame.ContainerWindow
    vistaInforme.setVisible(False)
    # Obtener el número de factura para ponerlo en el nombre del archivo
    sql = f'SELECT "FcNumero" FROM "FacturasColaborador" WHERE "FcId" = {fc_id}'
    rs = stat.executeQuery(sql)
    rs.first()
    numFactura = rs.getString(rs.findColumn('FcNumero'))

    archivo = uno.systemPathToFileUrl(dir_fac_colab + numFactura + '.pdf')
    # Imprimir la factura
    args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
    informe.storeToURL(archivo, args)
    informe.close(False)
    return


# ----------------------------------------------------------------------
# Imprime en pdf la factura con id fa_id
def imprimir_proforma(event=None):
    bas = CreateScriptService('Basic')
    form = event.Source.Model.Parent

    if form.getBoolean(form.findColumn("FaBloqueada")):
        bas.MsgBox('La factura ya está emitida.\nNo se puede generar proforma',
                   bas.MB_ICONEXCLAMATION, 'Error de facturación')
        return
    fa_id = form.getString(form.findColumn("FaId"))
    fa_concepto = form.getString(form.findColumn("FaConcepto"))

    sql = f'UPDATE "Filtros" SET "Valor" = {fa_id} WHERE "FiId" = 1'
    con = form.ActiveConnection
    stat = con.createStatement()
    stat.executeUpdate(sql)
    # abrir el informe y ocultarlo
    if fa_concepto:
        informe = bas.ThisDatabaseDocument.ReportDocuments \
            .getByName('FacturaProformaConcepto').open()
    else:
        informe = bas.ThisDatabaseDocument.ReportDocuments \
            .getByName('FacturaProforma').open()
    vistaInforme = informe.CurrentController.Frame.ContainerWindow
    vistaInforme.setVisible(False)
    # Nombre para el PDF creado
    archivo = uno.systemPathToFileUrl(dir_facturas + 'PRO-' + fa_id + '.pdf')
    # Crear el pdf, guardarlo y cerrar el informe
    args = (PropertyValue(Name='FilterName', Value='writer_pdf_Export'),)
    informe.storeToURL(archivo, args)
    informe.close(True)
    return


# ----------------------------------------------------------------------
# Ejecuta las rutinas necesarias para iniciar el programa
def iniciar_programa(event=None):
    abrir_menu_ppal(event)
    ocultar_base(event)
    cargar_config(event)
    return


# ----------------------------------------------------------------------
# Pone en blanco todos los campos de la tabla Filtros (para cancelar el filtrado)
def limpiar_filtros(event=None):
    bas = CreateScriptService('Basic')
    ds = bas.thisDatabaseDocument.DataSource
    con = ds.getConnection('', '')
    stat = con.createStatement()
    sql = """UPDATE "Filtros"  SET "Valor" = ''"""
    stat.executeQuery(sql)

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
    # DoCmd.SelectObject(acConstants.acDatabaseWindow)
    # DoCmd.Maximize()
    bas =CreateScriptService("Basic")
    vent = bas.ThisDatabaseDocument.CurrentController.Frame.ContainerWindow.IsMaximized = True
    return


# ----------------------------------------------------------------------
# Muestra el menu y barras de herramientas de un formulario que los tiene ocultos
def mostrar_menus(event=None):
    doc = event.Source
    frame = doc.CurrentController.Frame
    frame.LayoutManager.setVisible(True)
    return


# ----------------------------------------------------------------------
# Oculta Base. Se llama desde un formulario
def ocultar_base(event=None):
    bas =CreateScriptService("Basic")
    doc = bas.ThisDatabaseDocument
    doc.CurrentController.Frame.ContainerWindow.IsMinimized = True
    return


# ----------------------------------------------------------------------
# Esconde el menu y barras de herramientas de un formulario
def ocultar_menus(event=None):
    doc = event.Source
    frame = doc.CurrentController.Frame
    frame.LayoutManager.setVisible(False)



# ----------------------------------------------------------------------
# Reimprime una factura desde el formulario de facturas
def reimprimir_factura(event=None):
    bas = CreateScriptService('Basic')
    form = event.Source.Model.Parent
    bloq = form.getBoolean(form.findColumn("FaBloqueada"))
    if not bloq:
        mensaje('La factura no se puede imprimir porque no está generada',
                bas.MB_ICONINFORMATION, 'Error de impresión')
        return
    fa_id = form.getString(form.findColumn("FaId"))
    imprimir_factura(form, fa_id)
    return


# ----------------------------------------------------------------------
# Rutinas a ejecutar cuando se cierra el programa
def salir(event=None):
    bas = CreateScriptService('Basic')
    doc = CreateScriptService("SFDocuments.Document", bas.ThisDatabaseDocument)
    # TODO Comentar temporalmente la siguiente línea si se necesita trabajar en Base
    doc.RunCommand("CloseDoc")
    return


# ----------------------------------------------------------------------
# Mostrar XRay
def xray(objeto):
    bas = CreateScriptService("Basic")
    bas.Xray(objeto)


# ----------------------------------------------------------------------
def main(event=None):
    # from Xlib import display
    ui= CreateScriptService("UI")
    xray(ui)
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