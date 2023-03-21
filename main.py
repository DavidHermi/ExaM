import gi
import conexionBD
import DBConection

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, Gdk
from reportlab.platypus import Paragraph
from reportlab.platypus import SimpleDocTemplate
from reportlab.platypus import Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table


class Exame(Gtk.Window):
    def __init__(self):
        Gtk.Window.__init__(self, title="Exame 20-03-2023")
        caixaMain = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=4)
        caixaH1 = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)
        caixaH2 = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)
        caixaH3 = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)
        caixaH4 = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)
        caixaTV = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)

        self.bbdd = conexionBD.ConexionBD("modelosClasicos.dat")
        self.bbdd.conectaBD()
        self.bbdd.creaCursor()

        listComercialesVentas = Gtk.ListStore(int, str, str, str, str, str, str)
        self.filtrado_albaran = None
        filtro = listComercialesVentas.filter_new()
        filtro.set_visible_func(self.filtro_albaran)

        lblNome = Gtk.Label(label="Nome Comercial:")
        lblApelidos = Gtk.Label(label="Apelidos Comercial:")
        lblCodigo = Gtk.Label(label="Código oficina:")
        lblTelefono = Gtk.Label(label="Extensión teléfonica:")
        lblCargo = Gtk.Label(label="Cargo:")
        self.txtNome = Gtk.Entry()
        self.txtApelidos = Gtk.Entry()
        self.txtCodigo = Gtk.Entry()
        self.txtTelefono = Gtk.Entry()
        self.cmbCargo = Gtk.ComboBoxText()
        #self.cmbCargo.connect("changed", self.on_cmbCargo_changed, filtro, listComercialesVentas)
        caixaMain.pack_start(caixaH1, True, True, 0)

        nAlbaranes = self.bbdd.consultaSenParametros("select cargo from COMERCIALESVENTAS")
        dAlbaranes = self.bbdd.consultaSenParametros(
                "select * from COMERCIALESVENTAS where cargo in (select cargo from COMERCIALESVENTAS)")
        print(nAlbaranes)
        print(dAlbaranes)
        for albaran in nAlbaranes:
            self.cmbCargo.append_text(str(albaran[0]))
        self.cmbCargo.append_text(str(0))

        for dAlbaran in dAlbaranes:
            listComercialesVentas.append(dAlbaran)

        caixaH1.pack_start(lblNome, True, True, 2)
        caixaH1.pack_start(self.txtNome, True, True, 2)
        caixaH1.pack_start(lblApelidos, True, True, 2)
        caixaH1.pack_start(self.txtApelidos, True, True, 2)
        caixaH1.pack_start(lblCargo, True, True, 2)

        caixaH2.pack_start(lblCodigo, True, True, 2)
        caixaH2.pack_start(self.txtCodigo, True, True, 2)
        caixaH2.pack_start(lblTelefono, True, True, 2)
        caixaH2.pack_start(self.txtTelefono, True, True, 2)
        caixaH2.pack_start(self.cmbCargo, True, True, 2)

        caixaMain.pack_start(caixaH2, True, True, 0)
        caixaMain.pack_start(caixaH3, True, True, 0)
        caixaMain.pack_start(caixaH4, True, True, 0)

        lblDireccionCorreo = Gtk.Label(label="Dirección de correo")
        self.txtDireccionCorreo = Gtk.Entry()
        lblFormatoCorreo = Gtk.Label(label="Formato de correo")
        rbtHtml = Gtk.RadioButton(label="Html")
        rbtTextoPlano = Gtk.RadioButton(label="Texto plano")
        rbtPersonalizado = Gtk.RadioButton(label="Personalizado")

        caixaH3.pack_start(lblDireccionCorreo, True, True, 0)
        caixaH3.pack_start(self.txtDireccionCorreo, True, True, 0)

        caixaH4.pack_start(lblFormatoCorreo, True, True, 0)
        caixaH4.pack_start(rbtHtml, True, True, 0)
        caixaH4.pack_start(rbtTextoPlano, True, True, 0)
        caixaH4.pack_start(rbtPersonalizado, True, True, 0)

        self.trvComerciaisVentas = Gtk.TreeView(model=filtro)
        for i, titulo in enumerate(["codigo","nom comercial", "apelidos com", "Cod oficina", "Correoe","ext telefonica","Cargo"]):
            celda = Gtk.CellRendererText()
            columna = Gtk.TreeViewColumn(titulo, celda, text=i)
            self.trvComerciaisVentas.append_column(columna)

        caixaTV.pack_start(self.trvComerciaisVentas, True, True, 5)
        caixaMain.pack_start(caixaTV, True, True, 0)

        seleccion = self.trvComerciaisVentas.get_selection()
        seleccion.connect("changed", self.on_trvVistaProgramas_changed, self.txtNome, self.txtApelidos, self.txtCodigo, self.txtTelefono)

        self.btnEngadir = Gtk.Button(label="Engadir")
        self.btnEngadir.connect("clicked", self.on_button_clicked, listComercialesVentas)

        self.btnEditar = Gtk.Button(label="Editar")
        self.btnBorrar = Gtk.Button(label="Borrar")

        self.btnXerar = Gtk.Button(label="Xerar informe")
        self.btnXerar.connect("clicked", self.on_btnInforme_clicked, self.bbdd)
        caixaMain.pack_start(self.btnEngadir, True, True, 2)
        caixaMain.pack_start(self.btnXerar, True, True, 2)

        self.btnGardar = Gtk.Button(label="Gardar")
        self.btnDescartar = Gtk.Button(label="Descartar")

        self.add(caixaMain)
        self.connect("delete-event", Gtk.main_quit)
        self.show_all()

    def on_cmbCargo_changed(self, combo, filtro, modelo):
        seleccion = combo.get_active_iter()
        if seleccion is not None:
            cmbModelo = combo.get_model()
            cargos = cmbModelo[seleccion][0]
            self.filtrado_albaran = cargos
           # print("Cargo Seleccionado: %d" % self.filtrado_albaran)
        if seleccion is None:
            self.filtrado_albaran = None
        filtro.refilter()

    def filtro_albaran(self, modelo, fila, datos):
        if (self.filtrado_albaran is None or self.filtrado_albaran == 0):
            return True
        else:
            print("NumeroAlbaran: %d" % (modelo[fila][0] == self.filtrado_albaran))
            return modelo[fila][0] == self.filtrado_albaran
    def on_trvVistaProgramas_changed(self, selec, label, label2, label3, label4):
        modelo, punteiro = selec.get_selected()
        if punteiro is not None:
            self.txtNome.set_text(modelo [punteiro][1])
            self.txtApelidos.set_text(str(modelo [punteiro][2]))
            self.txtCodigo.set_text(modelo [punteiro][3])
            self.txtDireccionCorreo.set_text(modelo[punteiro][4])
            self.txtTelefono.set_text(modelo [punteiro][5])
    def on_button_clicked(self, boton, modelo):
        rexistro = (int(3),
                    self.txtNome.get_text(),
                    self.txtApelidos.get_text(),
                    self.txtCodigo.get_text(),
                    self.txtTelefono.get_text(),
                    self.txtDireccionCorreo.get_text(),
                    "Xefe ventas"
                    )
        modelo.append(rexistro)

    def on_btnInforme_clicked(self, boton, bd):
        doc = []
        hojaEstilo = getSampleStyleSheet()
        listaVenta = bd.consultaSenParametros("select * from COMERCIALESVENTAS")
        listaDetalles = bd.consultaSenParametros(
            "select * from COMERCIALESVENTAS where codigoComercial in(select codigoComercial from COMERCIALESVENTAS)")

        cabecera = hojaEstilo['Heading1']
        cabecera.pageBreakBefore = 0
        cabecera.keepWithNext = 1
        parrafo = Paragraph("Comerciales:", cabecera)
        doc.append(parrafo)
        doc.append(Spacer(0, 20))

        cuerpoTexto = hojaEstilo['BodyText']
        cuerpoTexto.keepWithNext = 0

        cabecera_al = hojaEstilo['Heading4']
        cuerpoTexto.keepWithNext = 0

        cadena = str()
        fila = []
        cabecera_tabla = ["nome", "Apelidos", "oficina","Correo", "telefono","Cargo"]
        ctabla = Table(cabecera_tabla)
        for venta in listaVenta:
            parrafo = Paragraph(
                "Codigo: " + str(venta[0]),
                cabecera_al)
            doc.append(parrafo)
            fila.append(cabecera_tabla)
            for detalles in listaDetalles:
                if venta[0] == detalles[0]:
                    fila.append(detalles[1:])

            tabla = Table(fila)
            doc.append(tabla)
            fila.clear()
            cadena = cadena + "\n"
            parrafo = Paragraph(cadena, cuerpoTexto)
            doc.append(parrafo)
            cadena = ""

        informe = SimpleDocTemplate("informe.pdf", pagesize=A4, showBoundary=0)
        informe.build(doc)


if __name__ == "__main__":
    Exame()
    Gtk.main()