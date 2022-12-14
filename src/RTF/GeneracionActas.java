/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package RTF;

import Excel.ControlArchivoExcel;
import Excel.gestorInformes;
import Utilidades.Fonts;
import com.lowagie.text.BadElementException;
import com.lowagie.text.Cell;
import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.field.*;
import com.lowagie.text.rtf.headerfooter.RtfHeaderFooter;
import com.sun.corba.se.impl.io.FVDCodeBaseImpl;
import java.awt.Color;
import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.*;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
 *
 * @author rjulio, rramos, kdelosreyes
 */
public class GeneracionActas {

    static String URL = "C:\\CIARP\\acta_.rtf";
    static String tipoProducto = "Ingreso a la Carrera Docente";
    static String cedula = "";
    static int bandera = 0;
    static int banderaTP = 0;
    DecimalFormat formateador = new DecimalFormat("#.#");
    static List<Map<String, String>> jerarquiaProducto = new ArrayList<>();
    static Map<String, String> listaDatosacta = new HashMap<>();
    static Map<String, String> Datosacta = new HashMap<>();
    static ArrayList<ArrayList<Map<String, String>>> DatosNumeralesActa = new ArrayList<ArrayList<Map<String, String>>>();
    public String ruta;
    public int indxgrado1 = 0;
    public int indxgrado2 = 0;
    public int indxgrado3 = 0;
    Map<String, String> respuestaerr = new HashMap<>();
    BaseFont arialFont;

    public GeneracionActas() {
        URL = "C:\\CIARP\\acta_.rtf";
        tipoProducto = "Ingreso a la Carrera Docente";
        cedula = "";
        bandera = 0;
        banderaTP = 0;
        jerarquiaProducto = new ArrayList<>();
        listaDatosacta = new HashMap<>();
        Datosacta = new HashMap<>();
        DatosNumeralesActa = new ArrayList<ArrayList<Map<String, String>>>();
        ruta = "";
        indxgrado1 = 0;
        indxgrado2 = 0;
        indxgrado3 = 0;
        

        InicializarJerarquia();

    }

    private void InicializarJerarquia() {
        Map<String, String> auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ingreso_a_la_Carrera_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ingreso a la Carrera Docente");
        auxiliarJerarquia.put("NORMA", "Art??culo 14 del Acuerdo Superior N?? 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ascenso_en_el_Escalafon_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ascenso en el Escalaf??n Docente");
        auxiliarJerarquia.put("NORMA", "Art??culo 27 del Acuerdo Superior N?? 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Titulacion");
        auxiliarJerarquia.put("NPRODUCTO", "Titulaci??n");
        auxiliarJerarquia.put("NORMA", "Art??culo 7 del Decreto 1279 delpp 2002, Art??culo Primero del Acuerdo 001 de 2004 del Grupo de Seguimiento al Decreto 1279 de 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Articulo");
        auxiliarJerarquia.put("NPRODUCTO", "Art??culo");
        auxiliarJerarquia.put("NPRODUCTOS", "Art??culos");
        auxiliarJerarquia.put("NORMA", "Literal a. numeral I, Art??culo 10 y literal a. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Video_Cinematograficas_o_Fonograficas");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n de Video Cinematogr??fica o Fonogr??fica");
        auxiliarJerarquia.put("NPRODUCTOS", "Producciones de Video Cinematogr??ficas o Fonogr??ficas");
        auxiliarJerarquia.put("NORMA", "Literal b. numeral I, Art??culo 10; literal b. numeral I, Art??culo 24; literal a. numeral I; literal a. numeral II, Articulo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Libro");
        auxiliarJerarquia.put("NPRODUCTOS", "Libros");
        auxiliarJerarquia.put("NORMA", "Literales c, d, e, Art??culo 10 y literales c, d, e. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Capitulo_de_Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Cap??tulo de Libro");
        auxiliarJerarquia.put("NPRODUCTOS", "Cap??tulos de Libro");
        auxiliarJerarquia.put("NORMA", "Literales c, d, e, Art??culo 10 y literales c, d, e. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Premios_Nacionales_e_Internacionales");
        auxiliarJerarquia.put("NPRODUCTO", "Premio Nacional o Internacional");
        auxiliarJerarquia.put("NPRODUCTOS", "Premios Nacionales o Internacionales");
        auxiliarJerarquia.put("NORMA", "Literal f. numeral I, Art??culo 10 y literal g. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Patente");
        auxiliarJerarquia.put("NPRODUCTO", "Patente");
        auxiliarJerarquia.put("NPRODUCTOS", "Patentes");
        auxiliarJerarquia.put("NORMA", "Literal g. numeral I, Art??culo 10 y literal h. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traduccion_de_Libros");
        auxiliarJerarquia.put("NPRODUCTO", "Traducci??n de Libro");
        auxiliarJerarquia.put("NPRODUCTOS", "Traducciones de Libro");
        auxiliarJerarquia.put("NORMA", "Literal h. numeral I, Art??culo 10 y literal f. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Obra_Artistica");
        auxiliarJerarquia.put("NPRODUCTO", "Obra Art??stica");
        auxiliarJerarquia.put("NPRODUCTOS", "Obras Art??sticas");
        auxiliarJerarquia.put("NORMA", "Literal i. numeral I, Art??culo 10; literal i. numeral I, Art??culo 24; literal b. numeral I; literal g. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_Tecnica");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n T??cnica");
        auxiliarJerarquia.put("NPRODUCTOS", "Producciones T??cnicas");
        auxiliarJerarquia.put("NORMA", "Literal j. numeral I, Art??culo 10 y literal j. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Software");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n de Software");
        auxiliarJerarquia.put("NPRODUCTOS", "Producciones de Software");
        auxiliarJerarquia.put("NORMA", "Literal k. numeral I, Art??culo 10 y literal k. numeral I, Art??culo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ponencias_en_Eventos_Especializados");
        auxiliarJerarquia.put("NPRODUCTO", "Ponencia en Evento Especializado");
        auxiliarJerarquia.put("NPRODUCTOS", "Ponencias en Eventos Especializados");
        auxiliarJerarquia.put("NORMA", "Literal c. numeral I, literal b. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Publicaciones_Impresas_Universitarias");
        auxiliarJerarquia.put("NPRODUCTO", "Publicaci??n Impresa Universitaria");
        auxiliarJerarquia.put("NPRODUCTOS", "Publicaciones Impresas Universitarias");
        auxiliarJerarquia.put("NORMA", "Literal d. numeral I, literal c. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Estudios_Posdoctorales");
        auxiliarJerarquia.put("NPRODUCTO", "Estudio Posdoctoral");
        auxiliarJerarquia.put("NPRODUCTOS", "Estudios Posdoctorales");
        auxiliarJerarquia.put("NORMA", "Literal e. numeral I, literal d. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Rese??as_Cr??ticas");
        auxiliarJerarquia.put("NPRODUCTO", "Rese??a Cr??tica");
        auxiliarJerarquia.put("NPRODUCTOS", "Rese??as Cr??ticas");
        auxiliarJerarquia.put("NORMA", "Literal f. numeral I, literal e. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traducciones");
        auxiliarJerarquia.put("NPRODUCTO", "Traducci??n");
        auxiliarJerarquia.put("NPRODUCTOS", "Traducciones");
        auxiliarJerarquia.put("NORMA", "Literal g. numeral I, literal f. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Direccion_de_Tesis");
        auxiliarJerarquia.put("NPRODUCTO", "Direcci??n de Tesis");
        auxiliarJerarquia.put("NPRODUCTOS", "Direcciones de Tesis");
        auxiliarJerarquia.put("NORMA", "Literal h. numeral I, literal h. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Evaluacion_como_par");
        auxiliarJerarquia.put("NPRODUCTO", "Evaluaci??n como par");
        auxiliarJerarquia.put("NPRODUCTOS", "Evaluaciones como par");
        auxiliarJerarquia.put("NORMA", "Literal i. numeral I, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);
    }

    public String Encode() {
        String cifrado = "" + System.currentTimeMillis();
        return cifrado;
    }

    public Map<String, String> GenerarActas(String ruta) throws DocumentException, IOException {
        this.ruta = ruta;
        ControlArchivoExcel con = new ControlArchivoExcel();
        List<Map<String, String>> listaDatos = new ArrayList<>();

        //<editor-fold defaultstate="collapsed" desc="Lectura Orden del d??a">
        String extP = ruta.substring(ruta.lastIndexOf(".") + 1);
        if (extP.equals("xlsx")) {
            listaDatos = con.LeerExcelDesdeAct(ruta, 2, "ORDEN DEL DIA");
            listaDatosacta = con.LeerExcelParametrosAct(ruta, 1, 1, "ORDEN DEL DIA");
        } else {
            listaDatos = con.LeerExcelDesde(ruta, 2, "ORDEN DEL DIA");
            listaDatosacta = con.LeerExcelParametros(ruta, 1, 1);
        }

//</editor-fold>
        ControlArchivo contArchivo = new ControlArchivo(ruta);
        contArchivo.LeerArchivo();
        BufferedReader br = contArchivo.getBuferDeLectura();
        String lineaDeTexto;

        Document documento = new Document(PageSize.LETTER);
        documento.setMargins(70, 70, 69, 55);

        String encode = Encode();

        URL = "C:\\CIARP\\acta_" + listaDatosacta.get("No_ACTA") + ".rtf";
        GeneracionDocumentoActa(listaDatos, documento);
        for (int j = 0; j < listaDatos.size(); j++) {

            for (Map.Entry<String, String> entry : listaDatos.get(j).entrySet()) {
                String key = entry.getKey();
                String value = entry.getValue();

            }

        }

        documento.close();
        int result = JOptionPane.showConfirmDialog(null, "??Desea abrir el documento?");
        if (result == JOptionPane.YES_OPTION) {
            Desktop.getDesktop().open(new File(URL));
        }
        gestorInformes gi = new gestorInformes(Datosacta, DatosNumeralesActa);
        gi.iniciar();

        return respuestaerr;
    }

    private Map<String, String> GeneracionDocumentoActa(List<Map<String, String>> listaDatos, Document documento) throws FileNotFoundException, BadElementException, IOException, DocumentException {
        arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
        RtfWriter2.getInstance(documento, new FileOutputStream(URL));

        documento.open();
        List<Map<String, String>> listaCorrespondencia = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Revision_de_la_correspondencia"});
        List<Map<String, String>> listaProposiciones = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Proposiciones_y_varios"});
        List<Map<String, String>> listaColciencias = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Art_Col"});
        //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">
        Fonts f = new Fonts(arialFont);
      
        Paragraph p = new Paragraph();
        int justificado = Paragraph.ALIGN_JUSTIFIED;
        int centrado = Paragraph.ALIGN_CENTER;

        //<editor-fold defaultstate="collapsed" desc="HEADER">
        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");

        Table headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("COMIT?? INTERNO DE ASIGNACI??N Y RECONOCIMIENTO DE PUNTAJE\n", Fonts.SetFontTwoStyle(Color.black, 11, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        //</editor-fold>

        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 200;//texto
        tam[1] = 9;// num page
        tam[2] = 2;//slide
        tam[3] = 9;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(100);

        footertable.setWidths(tam);

        try {
            celda = new Cell(new Paragraph("Acta N?? " + listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0), Fonts.SetFont(Color.black, 9, Fonts.BOLD)));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", "" + ex.getMessage());
            respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha del acta de la primera fila del documento");
            respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
        }
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setBorderWidthTop(2);
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 9, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setBorderWidthTop(2);
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("/", Fonts.SetFont(Color.black, 9, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setBorderWidthTop(2);
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 9, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        celda.setBorderWidthTop(2);
        footertable.addCell(celda);
        //</editor-fold>

        RtfHeaderFooter header = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(header);
        documento.setFooter(footer);
       
        //</editor-fold>

        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.BOLD));

        String[] a??o = listaDatosacta.get("FECHA_ACTA").split("/");
        p.add("\nACTA " + listaDatosacta.get("No_ACTA") + " DE " + a??o[2]);

        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFontTwoStyle(Color.black, 11, Fonts.BOLD));
        p.setAlignment(centrado);
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.setAlignment(justificado);
        try {
            p.add("En Santa Marta, a los " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 1) + ", se reunieron en sesi??n ordinaria los miembros del ");
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", "" + ex.getMessage());
            respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha del acta de la primera fila del documento");
            respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");

        }
        Chunk c = new Chunk("COMIT?? INTERNO DE ASIGNACI??N Y RECONOCIMIENTO DE PUNTAJE ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        String orden = "convocados por el Presidente de ??ste ??rgano colegiado, para tratar el siguiente orden del d??a:\n"
                + "\n"
                + "1. Verificaci??n del Qu??rum.\n"
                + "\n"
                + "2. Aprobaci??n del orden del d??a.\n"
                + "\n"
                + "3. Lectura y aprobaci??n del acta anterior.\n"
                + "\n";
        indxgrado1 = 4;
        if (listaCorrespondencia.size() > 0) {

            orden += (indxgrado1++) + ". Estudio de la Correspondencia\n"
                    + "\n";
        }

        orden += (indxgrado1++) + ". Solicitudes de los Docentes \n"
                + "\n";
        if (listaProposiciones.size() > 0) {

            orden += (indxgrado1++) + ". Proposiciones y varios\n"
                    + "\n";
        }

        orden += (indxgrado1++) + ". Cierre. \n\n";

        p.add(orden);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        c = new Chunk("DESARROLLO\n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);

        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.add("\n Se da inicio a la sesi??n, en el despacho del Vicerrector Acad??mico, siendo las " + listaDatosacta.get("HORA_INICIO") + "\n \n");
        c = new Chunk("1. VERIFICACI??N DEL QU??RUM ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        p.add("se verifica el qu??rum para deliberar contando con los siguientes asistentes:\n"
                + "\n"
                + "1. Oscar Humberto Garcia Vargas - Presidente del Comit??\n"
                + "2. Jorge Enrique Elias Caro ??? Vicerrector de Investigaci??n\n"
                + "3. Haidy Rocio Oviedo Cordoba (Representante de los docentes del ??rea Ciencias de la Salud)\n"
                + "4. Rocio Del Pilar Garcia Urue??a (Representante de los docentes del ??rea de Matem??ticas y Ciencias Naturales)\n"
                + "5. Rolando Enrique Escorcia Caballero (Representante de los docentes del ??rea de Ciencias de la Educaci??n)\n"
                + "6. Marco Francisco Gaviria Rueda (Representante de los docentes del ??rea de Arquitectura, Bellas Artes y Afines)\n"
                + "7. Karen Gishelle Buelvas Ferreira (Profesional Especializado de la Vicerrector??a Acad??mica)\n"
                + "8. Kennys De Los Reyes Castillo (Profesional Universitario de la Vicerrector??a Acad??mica)\n"
                + "9. Diana Paola Orozco Tete (Contratista)\n"
                + "\n"
                + "Considerando lo establecido en el art??culo 5to del Acuerdo Superior N?? 021 de 2009, que dispone la composici??n y requisitos de los miembros del Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje atendiendo los principios de eficacia, econom??a y celeridad que enmarcan todas las actuaciones administrativas, ??ste ??rgano colegiado, decide por unanimidad designar a la Profesional Especializado de la Vicerrector??a Acad??mica Karen Gishelle Buelvas Ferreira, como secretaria t??cnica para que se encargue de la parte operativa del CIARP.  \n"
                + "\n");
        c = new Chunk("2. APROBACI??N DEL ORDEN DEL D??A ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        p.add("se aprueba el orden del d??a.\n"
                + "\n");
        c = new Chunk("3. LECTURA Y APROBACI??N DEL ACTA ANTERIOR ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        try {
            p.add("le??da el Acta N?? " + listaDatosacta.get("ACTA_ANTERIOR") + " de fecha " + fechaEnletras(listaDatosacta.get("FECHA_ACTA_ANTERIOR"), 0) + " es aprobada por unanimidad por los miembros del Comit??.  \n"
                    + "\n");
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", "" + ex.getMessage());
            respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha del acta anterior de la primera fila del documento");
            respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
        }
        indxgrado1 = 3;

        documento.add(p);
        Map<String, String> datos1 = new HashMap<>();
        ArrayList<Map<String, String>> datos2 = new ArrayList<>();

        if (listaCorrespondencia.size() > 0) {
            indxgrado1++;
            p = new Paragraph(10);
            c = new Chunk(indxgrado1 + ". ESTUDIO DE LA CORRESPONDENCIA: \n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
            p.add(c);
            documento.add(p);
            indxgrado2 = 1;
            for (int l = 0; l < listaCorrespondencia.size(); l++) {

                p = new Paragraph(10);
                p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                p.setAlignment(justificado);
                String numeral = indxgrado1 + "." + (indxgrado2++);

                //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                datos1 = new HashMap<>();
                datos1.put("TIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("TIPO_PRODUCTO")));
                datos1.put("SUBTIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("SUBTIPO_PRODUCTO")));
                datos1.put("IDDOCENTE", "" + listaCorrespondencia.get(l).get("No._IDENTIFICACION"));
                datos1.put("DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_DOCENTE")));
                datos1.put("NOMBRE_SOLICITUD", "" + Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_SOLICITUD")));
                datos1.put("RESPUESTA_CIARP", "" + Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("RESPUESTA_CIARP")));
                datos1.put("NUMERAL", "" + numeral);
                datos1.put("POS", "1");
                datos2.add(datos1);
                //</editor-fold>

                c = new Chunk(numeral + ". ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_SOLICITUD"))) + " \n");
                c = new Chunk("Decisi??n: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("DECISION"))) + ".\n");
                documento.add(p);
            }

        }

        indxgrado1++;

        p = new Paragraph(10);
        c = new Chunk(indxgrado1 + ". ESTUDIO DE LAS SOLICITUDES DE DOCENTES: \n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        documento.add(p);
        indxgrado2 = 0;
        for (int i = 0; i < jerarquiaProducto.size(); i++) {
            List<Map<String, String>> listadatosxTipoProducto = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + jerarquiaProducto.get(i).get("PRODUCTO")});
            if (listadatosxTipoProducto.size() > 0) {
                Datosacta.put("NOMBRE_ARCHIVO", "numerales_acta_");
                Datosacta.put("NUMACTA", listadatosxTipoProducto.get(0).get("ACTA"));
                indxgrado2++;
                p = new Paragraph(10);
                p.setFont(Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.setAlignment(justificado);
                p.add(indxgrado1 + "." + indxgrado2 + ". Estudio de solicitudes por " + jerarquiaProducto.get(i).get("NPRODUCTO") + ":\n");
                documento.add(p);

                List<Map<String, String>> listadocentexTipoProducto = data_list(1, listadatosxTipoProducto, new String[]{"No._IDENTIFICACION"});
                indxgrado3 = 0;
                for (int j = 0; j < listadocentexTipoProducto.size(); j++) {
                    List<Map<String, String>> listadatosdocentexTipoProducto = data_list(3, listadatosxTipoProducto, new String[]{"No._IDENTIFICACION<->" + listadocentexTipoProducto.get(j).get("No._IDENTIFICACION")});
                    indxgrado3 += 1;
                    p = new Paragraph(10);
                    p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                    p.setAlignment(justificado);
                    String numeral = indxgrado1 + "." + indxgrado2 + "." + indxgrado3;
                    c = new Chunk(numeral + ". " + Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOMBRE_DOCENTE")) + ":\n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);
                    c = new Chunk("Identificaci??n: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);
                    p.add(FormatoCedula(listadocentexTipoProducto.get(j).get("No._IDENTIFICACION")) + "\n" + Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("FACULTAD")) + ".\n");
                    c = new Chunk("Categor??a: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);
                    try {
                        p.add(listadocentexTipoProducto.get(j).get("CATEGORIA_DOCENTE") + " - " + listadocentexTipoProducto.get(j).get("TIPO_VINCULACION") + " desde el " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("FECHA_INGRESO")), 0) + ".\n");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                        respuestaerr.put("ESTADO", "ERROR");
                        respuestaerr.put("MENSAJE", "" + ex.getMessage());
                        respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha de vinculaci??n del docente " + Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOMBRE_DOCENTE")));
                        respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
                    }
                    c = new Chunk("Tipo de solicitud: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);

                    p.add((Utilidades.Utilidades.decodificarElemento(listadatosxTipoProducto.get(0).get("TIPO_PRODUCTO")).equals("Ascenso_en_el_Escalafon_Docente") || Utilidades.Utilidades.decodificarElemento(listadatosxTipoProducto.get(0).get("TIPO_PRODUCTO")).equals("Ingreso_a_la_Carrera_Docente")
                            ? "Puntos por" : "Puntos por la publicaci??n de " + listadatosdocentexTipoProducto.size()) + " "
                            + (listadatosdocentexTipoProducto.size() > 1
                            ? jerarquiaProducto.get(i).get("NPRODUCTOS")
                            : jerarquiaProducto.get(i).get("NPRODUCTO"))
                            + ". (" + jerarquiaProducto.get(i).get("NORMA") + ").\n");

                    for (int k = 0; k < listadatosdocentexTipoProducto.size(); k++) {
                        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                        datos1 = new HashMap<>();
                        datos1.put("TIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO")));
                        datos1.put("SUBTIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SUBTIPO_PRODUCTO")));
                        datos1.put("IDDOCENTE", "" + listadatosdocentexTipoProducto.get(k).get("No._IDENTIFICACION"));
                        datos1.put("DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_DOCENTE")));
                        datos1.put("NOMBRE_SOLICITUD", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD")));
                        datos1.put("RESPUESTA_CIARP", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("RESPUESTA_CIARP")));
                        datos1.put("NUMERAL", "" + numeral);
                        datos1.put("POS", "" + (k + 1));
                        datos2.add(datos1);
                        //</editor-fold>

                        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                        p.setAlignment(justificado);
                        c = new Chunk("Soporte" + (listadatosdocentexTipoProducto.size() > 1
                                ? (" " + getNombreNumero(k + 1, jerarquiaProducto.get(i).get("ARTICULO")) + " " + jerarquiaProducto.get(i).get("NPRODUCTO") + ": ")
                                : ": "),
                                 Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                        p.add(c);
                        String soportes = getSoportes(listadatosdocentexTipoProducto, k);
                        p.add(soportes + "\n");

                    }

                    String decision;
                    try {
                        decision = getDecision(listadatosdocentexTipoProducto);
                        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                        p.setAlignment(justificado);
                        c = new Chunk("Decisi??n: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                        p.add(c);
                        p.add("Revisada la documentaci??n y despu??s de analizar lo establecido en las normas, el Comit?? decide " + decision + "\n");

                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                        respuestaerr.put("ESTADO", "ERROR");
                        respuestaerr.put("MENSAJE", "" + ex.getMessage());
                        System.out.println(" ERROROOO " + ex.getMessage());
                        respuestaerr.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(0).get("NOMBRE_DOCENTE")));
                        respuestaerr.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO")));
                        if (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD")).length() <= 100) {
                            respuestaerr.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD")));
                        } else {
                            respuestaerr.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD")).substring(0, 100));
                        }
                        return respuestaerr;
                    }

                    if (!Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOTA")).equals("N/A")) {
                        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                        p.setAlignment(justificado);
                        c = new Chunk("Nota: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                        p.add(c);
                        p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOTA"))) + "\n");

                    }

                    documento.add(p);

                }
            }
        }

        if (listaColciencias.size() > 0) {
            List<Map<String, String>> listadocentexArticuloColciencias = data_list(1, listaColciencias, new String[]{"No._IDENTIFICACION"});
            indxgrado1++;
            p = new Paragraph(10);
            c = new Chunk(indxgrado1 + ". REVISI??N Y ASIGNACI??N DE PUNTOS A LOS DOCENTES CUYOS ART??CULOS QUEDARON PENDIENTES POR HOMOLOGACI??N DE COLCIENCIAS PARA EL A??O 2019: \n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
            p.add(c);
            documento.add(p);
            indxgrado2 = 1;
            p = new Paragraph(10);
            p.add("Teniendo en cuenta la respuesta brindada por el Grupo de Seguimiento al R??gimen Salarial y Prestacional de los Docentes de universidades P??blicas a la consulta: "
                    + "??Se pueden reconocer puntos salariales que sean solicitados y estudiados antes de la actualizaci??n de la base de datos de Colciencias"
                    + " y hacer el pago de estos una vez se conozca la categor??a en que se encuentra la revista; y de igual manera pagarse los meses transcurridos"
                    + " a partir del estudio de la solicitud?\n");
            p = new Paragraph(10);
            c = new Chunk("RESPUESTA:", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
            p.add(c);
            p = new Paragraph(10);
            p.add("???En este caso se reconoce provisionalmente el puntaje hasta tanto se publique por Colciencias por la clasificaci??n de las revistas, se aclara que no hay retroactividad,"
                    + " y solo se asigna el puntaje hasta tanto que se encuentre en firme la clasificaci??n de las revistas de cada instituci??n, se deber?? tener en cuenta que la Homologaci??n "
                    + "se surte en el ??ltimo trimestre de cada a??o.\n"
                    + "As?? las cosas se har?? efectivo el reconocimiento solo a partir del momento en el que se haga p??blico el listado"
                    + " de revistas extranjeras homologadas para la respectiva vigencia??? \n"
                    + " \n"
                    + "Y que el Comit?? se acogi?? a esta respuesta; se hace asignaci??n de los puntos salariales que hab??an quedado pendientes "
                    + "a los docentes que solicitaron puntos por los art??culos publicados en las revistas que no se encontraban en la actualizaci??n de"
                    + " base de datos de Colciencias, seg??n la siguiente relaci??n: \n");

            for (int l = 0; l < listadocentexArticuloColciencias.size(); l++) {
                List<Map<String, String>> listainfoDocentesColciencias = data_list(3, listaColciencias, new String[]{"No._IDENTIFICACION<->" + listadocentexArticuloColciencias.get(l).get("No._IDENTIFICACION")});
                p = new Paragraph(10);
                p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                p.setAlignment(justificado);
                String numeral = indxgrado1 + "." + (indxgrado2++);
                c = new Chunk(numeral + ". DOCENTE: " + listadocentexArticuloColciencias.get(l).get("NOMBRE_DOCENTE") + "\n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.add(c);

                for (int k = 0; k < listainfoDocentesColciencias.size(); k++) {
                    //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                    datos1 = new HashMap<>();
                    datos1.put("TIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("TIPO_PRODUCTO")));
                    datos1.put("SUBTIPO_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("SUBTIPO_PRODUCTO")));
                    datos1.put("IDDOCENTE", "" + listainfoDocentesColciencias.get(k).get("No._IDENTIFICACION"));
                    datos1.put("DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("NOMBRE_DOCENTE")));
                    datos1.put("NOMBRE_SOLICITUD", "" + Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("NOMBRE_SOLICITUD")));
                    datos1.put("RESPUESTA_CIARP", "" + Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("RESPUESTA_CIARP")));
                    datos1.put("NUMERAL", "" + numeral);
                    datos1.put("POS", "" + (k + 1));
                    datos2.add(datos1);
                    //</editor-fold>

                    p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                    p.setAlignment(justificado);
                    c = new Chunk("Art??culo " + (listainfoDocentesColciencias.size() > 1 ? "N?? " + (k + 1) + ": " : ": "), Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);
                    String soportes = getSoportesColciencias(listainfoDocentesColciencias, k);
                    p.add(soportes + "");

                }

                c = new Chunk("DECISI??N: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.add(c);

                try {
                    String decision;
                    decision = getDecisionColciencias(listainfoDocentesColciencias);
                    p.add(decision + "\n");
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                    respuestaerr.put("ESTADO", "ERROR");
                    respuestaerr.put("MENSAJE", "" + ex.getMessage());
                    respuestaerr.put("LINEA_ERROR_DOCENTE", "" + listainfoDocentesColciencias.get(l).get("NOMBRE_DOCENTE"));
                    respuestaerr.put("LINEA_ERROR_PRODUCTO", "" + listainfoDocentesColciencias.get(l).get("TIPO_PRODUCTO"));
                    if (listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD").length() <= 100) {
                        respuestaerr.put("NOMBRE_PRODUCTO", "" + listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD"));
                    } else {
                        respuestaerr.put("NOMBRE_PRODUCTO", "" + listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD").substring(0, 100));
                    }
                    return respuestaerr;
                }

                documento.add(p);

                if (!Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(l).get("NOTA")).equals("N/A")) {
                    p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
                    p.setAlignment(justificado);
                    c = new Chunk("Nota: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                    p.add(c);
                    p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(l).get("NOTA"))) + "\n");

                }
            }

        }

        if (listaProposiciones.size() > 0) {
            indxgrado1++;
            indxgrado2 = 0;
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
            p.setAlignment(justificado);
            c = new Chunk(indxgrado1 + ". PROPOSICIONES Y VARIOS: \n\n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
            p.add(c);
            for (int y = 0; y < listaProposiciones.size(); y++) {
                indxgrado2++;
                p.add(indxgrado1 + "." + indxgrado2 + " " + ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaProposiciones.get(y).get("NOMBRE_SOLICITUD"))) + "\n");
                c = new Chunk("Decisi??n: ", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaProposiciones.get(y).get("DECISION"))) + "\n");
            }

            documento.add(p);
        }
        DatosNumeralesActa.add(datos2);
        indxgrado1++;
        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.setAlignment(justificado);
        c = new Chunk(indxgrado1 + ". CIERRE: \n\n", Fonts.SetFont(Color.black, 11, Fonts.BOLD));
        p.add(c);
        p.add(" Siendo las " + listaDatosacta.get("HORA_FIN") + " se da por terminada la sesi??n");
        documento.add(p);
        respuestaerr.put("ESTADO", "OK");
        respuestaerr.put("MENSAJE", "Las cartas se generaron satisfactoriamente.");
//
        return respuestaerr;
    }

    private List<Map<String, String>> data_list(int caso, List<Map<String, String>> lista, String[] datos) {
        List<Map<String, String>> rlista = new ArrayList<Map<String, String>>();
        try {
            switch (caso) {
                case 1: {//para listar los datos por el dato enviado
                    for (Map<String, String> lis : lista) {
                        boolean encontro = false;
                        for (Map<String, String> lr : rlista) {
                            if (lis.get(datos[0]).equals(lr.get(datos[0]))) {
                                encontro = true;
                                break;
                            }
                        }
                        if (!encontro) {
                            rlista.add(lis);
                        }
                    }
                    break;
                }
                case 10: {
                    for (Map<String, String> lis : lista) {
                        boolean encontro = false, enc = true;
                        for (Map<String, String> lr : rlista) {
                            boolean cond = false;
                            for (int i = 0; i < datos.length; i++) {
                                if (lis.get(datos[i]).equals("")) {
                                    cond = true;
                                    break;
                                }
                                if (i == 0) {
                                    cond = lis.get(datos[i]).equals(lr.get(datos[i]));
                                } else {
                                    cond = cond && lis.get(datos[i]).equals(lr.get(datos[i]));
                                }
                            }
                            if (cond) {
                                encontro = true;
                                break;
                            }
                        }
                        if (!encontro) {
                            rlista.add(lis);
                        }
                    }
                    break;
                }
                case 2: {////para listar datos por la key mandada y el valor mandado
                    for (Map<String, String> lis : lista) {
                        if (lis.get(datos[0]).equals(datos[1])) {
                            rlista.add(lis);
                        }
                    }
                    break;
                }

                case 3: { //para listar los datos por los datos enviados de de la siguiente forma
                    //k<->val, k<->val                    
                    for (Map<String, String> lis : lista) {
                        int coincidencias = 0;
                        for (String prm : datos) {
                            String[] item = prm.split("<->");
                            if (lis.get(item[0]).equals(item[1])) {
                                coincidencias++;
                            }
                        }
                        if (coincidencias == datos.length) {
                            rlista.add(lis);
                        }
                    }
                    break;
                }

            }

        } catch (Exception e) {
        }
        return rlista;
    }

    private List<Map<String, String>> data_list(int caso, List<Map<String, String>> lista, String[] datos, String[] datos2) {
        List<Map<String, String>> rlista = new ArrayList<Map<String, String>>();
        try {
            switch (caso) {  //////////////LISTAR DATOS POR VECTOR dE COiNCIDENCIAS              
                case 1: {
                    for (Map<String, String> lis : lista) {
                        int coincidencias = 0;
                        for (String prm : datos2) {
                            String[] item = prm.split("<->");
                            if (lis.get(item[0]).equals(item[1])) {
                                coincidencias++;
                            }
                        }
                        if (coincidencias == datos2.length) {
                            boolean encontro = false;
                            for (Map<String, String> lr : rlista) {

                                if (lis.get(datos[0]).equals(lr.get(datos[0])) || lis.get(datos[0]).trim().equals("")) {
                                    encontro = true;
                                    break;
                                }
                            }
                            if (!encontro && !lis.get(datos[0]).trim().equals("")) {
                                rlista.add(lis);
                            }
                        }
                    }
                    break;
                }
                case 10: {
                    for (Map<String, String> lis : lista) {
                        int coincidencias = 0;
                        for (String prm : datos2) {
                            String[] item = prm.split("<->");
                            if (lis.get(item[0]).equals(item[1])) {
                                coincidencias++;
                            }
                        }
                        if (coincidencias == datos2.length) {
                            boolean encontro = false, enc = true;
                            for (Map<String, String> lr : rlista) {
                                boolean cond = false;
                                for (int i = 0; i < datos.length; i++) {
                                    if (lis.get(datos[i]).equals("")) {
                                        cond = true;
                                        break;
                                    }
                                    if (i == 0) {
                                        cond = lis.get(datos[i]).equals(lr.get(datos[i]));
                                    } else {
                                        cond = cond && lis.get(datos[i]).equals(lr.get(datos[i]));
                                    }
                                }
                                if (cond) {
                                    encontro = true;
                                    break;
                                }
                            }
                            if (!encontro) {
                                rlista.add(lis);
                            }
                        }
                    }
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();

        }
        return rlista;
    }

    private String getDecision(List<Map<String, String>> listadatosdocentexTipoProducto) throws Exception {
        String respuesta = "";
        String respuestaxEstado = "";
        int banderasuma = 0;
        double sumatoria_puntos = 0;
        String articulo = "";
        for (int f = 0; f < listadatosdocentexTipoProducto.size(); f++) {
            try {
                sumatoria_puntos += Double.parseDouble(ValidarNumero(listadatosdocentexTipoProducto.get(f).get("PUNTOS").replace(",", ".")));
            } catch (Exception ex) {
                Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                throw new Exception(" " + ex.getMessage());
            }
            if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Aprobado")) {
                banderasuma = 1;
            }

            articulo = getArticuloxTipoProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"));
            respuestaxEstado += (listadatosdocentexTipoProducto.size() > 1 ? " por " + articulo + " "
                    + (getNombreNumero((f + 1), getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "ARTICULO")) + " "
                    + getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO") + " ") : "");
            if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Aprobado")) {
                if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {

                    respuestaxEstado += "aprobar la promoci??n en el escalaf??n docente a la categor??a "
                            + listadatosdocentexTipoProducto.get(f).get("NOMBRE_SOLICITUD")
                            + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")
                            ? " al" : " a la") + " "
                            + "docente de planta "
                            + listadatosdocentexTipoProducto.get(f).get("NOMBRE_DOCENTE")
                            + " y asignar " + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + " ("
                            + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") " + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE");

                    if (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals("N/A")) {
                        respuestaxEstado += " a partir de la fecha de la presente sesi??n";
                    } else if (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).length() > 10) {
                        respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")) + ".";
                    } else {
                        respuestaxEstado += " a partir de " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")), 0) + "";
                    }

                } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Ingreso_a_la_Carrera_Docente")) {
                    respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Titulacion")) {
                    respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                } else {

                    boolean cond = listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").equals("puntos salariales");
                    if (cond) {

                        respuestaxEstado += "asignarle "
                                + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")
                                ? "al" : "a la") + " "
                                + "docente "
                                + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") "
                                + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE");

                        if (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals("N/A")) {
                            respuestaxEstado += " a partir de la fecha de la presente sesi??n";
                        } else if (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).length() > 10) {
                            respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")) + ".";
                        } else {
                            respuestaxEstado += " a partir del " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")), 0) + "";
                        }

                        if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/A")) {
                            respuestaxEstado += " considerando que " + articulo + " "
                                    + getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                                    

                                    + " corresponde a un(a) "
                                    + (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")).equals("N/A")
                                    ? "producto"
                                    : Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")) + " ")
                                   
                                    + (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                    ? " de car??cter " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL"))
                                    : "")
                                    + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")) + ") ";
                        }
                        try {
                            if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) > 3) {
                                if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) < 6) {
                                    respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal b; numeral III, art??culo 10 del Decreto 1279 de 2002).";
                                } else {
                                    respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal c; numeral III, art??culo 10 del Decreto 1279 de 2002).";
                                }
                            }
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(" " + ex.getMessage());
                        }

                    } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                        respuestaxEstado += "reconocer "
                                + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")
                                ? "al" : "a la") + " "
                                + "docente "
                                + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") " + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").replace("puntos de bonificacion", "puntos de bonificaci??n") + ".";

                        if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/A")) {
                            respuestaxEstado += " considerando que " + articulo + " "
                                    + getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                                    + " corresponde a un(a) "
                                    + (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")).equals("N/A")
                                    ? "producto"
                                    : Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")) + " ")
                                    + (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                    ? " de car??cter " + listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")
                                    : "")
                                    + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")) + ") ";
                        }
                        try {
                            if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) > 3) {
                                if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) < 6) {
                                    respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal b; numeral I, art??culo 21 del Decreto 1279 de 2002).";
                                } else {
                                    respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal c; numeral I, art??culo 21 del Decreto 1279 de 2002).";
                                }
                            }
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(" " + ex.getMessage());
                        }

                    } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").trim().equals("no aplica") || listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").trim().equals("convalidacion")) {
                        respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                    }
                }
            } else if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Rechazado")) {
                respuestaxEstado += "no dar tr??mite a la solicitud "
                        + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")
                        ? "del" : "de la") + " "
                        + "docente en raz??n a que " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
            } else if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Enviar a pares")) {
                respuestaxEstado += "enviar el producto a revisi??n por parte de pares externos de Colciencias teniendo en cuenta lo establecido en el Art??culo 15 del Decreto 1279 del 2002";
            } else {
                respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
            }

        }

        String add = (listadatosdocentexTipoProducto.size() > 1
                ? "\nPara un total de " + getNumeroDecimal("" + formateador.format(sumatoria_puntos)) + " (" + formateador.format(sumatoria_puntos) + ") "
                + listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE")
                + " por la productividad presentada"
                : "");

        respuesta = respuestaxEstado;
        if (banderasuma == 1) {
            respuesta += ((!listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon _Docente")
                    && !listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Titulacion")
                    && listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE").equals("puntos salariales"))
                    ? add + ", de conformidad con lo establecido en el Numeral 22 del Art??culo Primero del Acuerdo de Seguimiento N?? 001 de 2004 y al Par??grafo III del Art??culo 12 del Decreto 1279 de 2002"
                    : listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE").equals("puntos de bonificacion") ? add : "");
        }

        return respuesta;
    }

    private String getSoportes(List<Map<String, String>> listadatosdocentexTipoProducto, int k) {
        String respuestaSoporte = "";
        String datosProducto = "";
        String datosSoporte = "";
        String repl = "";

        switch (listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO")) {
            case "Ingreso_a_la_Carrera_Docente":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                respuestaSoporte = datosSoporte;
                break;
            case "Ascenso_en_el_Escalafon_Docente":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                respuestaSoporte = datosSoporte;
                break;
            case "Articulo":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de la " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia del art??culo" + datosProducto + "; " + datosSoporte;
                break;
            case "Libro":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia del libro" + datosProducto + "; " + datosSoporte;
                break;
            case "Capitulo_de_Libro":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia del capitulo" + datosProducto + "; " + datosSoporte;
                break;
            case "Ponencias_en_Eventos_Especializados":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; en el " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia de la ponencia" + datosProducto + "; " + datosSoporte;
                break;
            case "Publicaciones_Impresas_Universitarias":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia de la publicaci??n impresa universitaria" + datosProducto + "; " + datosSoporte;
                break;
            case "Rese??as_Cr??ticas":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de la " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia de la rese??a" + datosProducto + "; " + datosSoporte;
                break;
            case "Traducciones":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                respuestaSoporte = "Copia de la traducci??n" + datosProducto + "; " + datosSoporte;
                break;
            case "Direccion_de_Tesis":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD")) + "\" ";

                respuestaSoporte = "Copia del acta de sustentaci??n de la" + datosProducto + "; " + datosSoporte;
                break;
            default:
                if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")).equals("N/A")) {
                    datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                            + "\" " + ValidarNumeroDec(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")) + " autor(es)";

                    respuestaSoporte = "Copia de " + listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO") + datosProducto + ", " + datosSoporte;
                } else {
                    datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                    datosProducto = " " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                            + " ";

                    respuestaSoporte = "Copia de" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO")) + datosProducto + ", " + datosSoporte;

                }
                break;
        }

        return respuestaSoporte;
    }

    private String getNumeroDecimal(String numero) {
        String retorno = "";

        if (numero.indexOf(",") > -1) {

            numero = numero.replace(",", ".");
        }

        if (numero.indexOf(".") > -1) {
            System.out.println("numero------>" + numero);
            Double dat = Double.parseDouble(numero);
            System.out.println("dat------>" + dat);
            DecimalFormat df = new DecimalFormat("#.0");
            System.out.println("df.format(dat)------>" + df.format(dat));
            numero = df.format(dat);
            System.out.println("numero------>" + numero);
            numero = numero.replace(".", ",");
            System.out.println("numero final------>" + numero);
        }

        if (!numero.equals("N/A")) {
            if (numero.indexOf(",") > -1) {
                String[] numrs = numero.replace(",", "::").split("::");
                retorno = numeroEnLetras(Integer.parseInt(numrs[0].equals("") ? "0" : numrs[0]));
                if (Integer.parseInt(numrs[1]) > 0) {
                    retorno += " coma ";
                    retorno += numeroEnLetras(Integer.parseInt(numrs[1]));
                }
            } else {
                retorno = numeroEnLetras(Integer.parseInt(numero));
            }
        }

        return retorno;
    }

    private String numeroEnLetras(int numero) {
        String[] Unidades, Decenas, Centenas;
        String Resultado = "";

        /**
         * ************************************************
         * Nombre de los n??meros
         * ************************************************
         */
        Unidades = new String[]{"", "Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Diecis??is", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veinti??n", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve"};
        Decenas = new String[]{"", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa", "Cien"};
        Centenas = new String[]{"", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos"};

        if (numero == 0) {
            Resultado = "Cero";
        } else if (numero >= 1 && numero <= 29) {
            Resultado = Unidades[numero];
        } else if (numero >= 30 && numero <= 100) {
            String agregado = "";
            if (numero % 10 != 0) {
                agregado = " y " + numeroEnLetras(numero % 10);
            } else {
                agregado = "";
            }
            Resultado = Decenas[numero / 10] + agregado;
        } else if (numero >= 101 && numero <= 999) {
            String agregado = "";
            if (numero % 100 != 0) {
                agregado = " " + numeroEnLetras(numero % 100);
            } else {
                agregado = "";
            }
            Resultado = Centenas[numero / 100] + agregado;
        } else if (numero >= 1000 && numero <= 1999) {
            String agregado = "";
            if (numero % 1000 != 0) {
                agregado = " " + numeroEnLetras(numero % 1000);
            } else {
                agregado = "";
            }
            Resultado = "Mil" + agregado;
        } else if (numero >= 2000 && numero <= 999999) {
            String agregado = "";
            if (numero % 1000 != 0) {
                agregado = " " + numeroEnLetras(numero % 1000);
            } else {
                agregado = "";
            }
            Resultado = numeroEnLetras(numero / 1000) + " Mil" + agregado;
        } else if (numero >= 1000000 && numero <= 1999999) {
            String agregado = "";
            if (numero % 1000000 != 0) {
                agregado = " " + numeroEnLetras(numero % 1000000);
            } else {
                agregado = "";
            }
            Resultado = "Un Mill??n" + agregado;
        } else if (numero >= 2000000 && numero <= 1999999999) {
            String agregado = "";
            if (numero % 1000000 != 0) {
                agregado = " " + numeroEnLetras(numero % 1000000);
            } else {
                agregado = "";
            }
            Resultado = numeroEnLetras(numero / 1000000) + " Millones" + agregado;
        }
        return Resultado.toLowerCase();
    }

    private String numeroOrdinales(int numero) {
        String[] Unidades, Decenas, Centenas;
        String Resultado = "";

        /**
         * ************************************************
         * Nombre de los n??meros
         * ************************************************
         */
        Unidades = new String[]{"", "Primero", "Segundo", "Tercero", "Cuarto", "Quinto", "Sexto", "S??ptimo", "Octavo", "Noveno", "D??cimo", "Und??cimo", "Duod??cimo"};
        Decenas = new String[]{"", "Decimo", "Vig??simo", "Trig??simo", "Cuadrag??simo", "Quincuag??simo", "Sexag??simo", "Septuag??simo", "Octog??simo", "Nonag??simo"};
        Centenas = new String[]{"", "Cent??simo", "Ducent??simo", "Tricent??simo", "Cuadringent??simo", "Quingent??simo", "Sexcent??simo", "Septingent??simo", "Octingent??simo", "Noningent??simo"};

        if (numero == 0) {
            Resultado = "Cero";
        } else if (numero >= 1 && numero <= 12) {
            Resultado = Unidades[numero];
        } else if (numero >= 13 && numero <= 100) {
            String agregado = "";
            if (numero % 10 != 0) {
                agregado = "" + numeroOrdinales(numero % 10);
            } else {
                agregado = "";
            }
            Resultado = Decenas[numero / 10] + agregado;
        } else if (numero >= 101 && numero <= 999) {
            String agregado = "";
            if (numero % 100 != 0) {
                agregado = " " + numeroOrdinales(numero % 100);
            } else {
                agregado = "";
            }
            Resultado = Centenas[numero / 100] + agregado;

        }
        return Resultado.toLowerCase();
    }

    private String ValidarNumero(String numero) throws Exception {
        return (numero.equals("N/A") ? "0" : numero);
    }

    private String fechaEnletras(String fecha, int opc) throws Exception {// 7/08/2012

        String fechaletra = "";
        if (!fecha.equals("N/A")) {
            try {
                String[] dividirFecha = fecha.split("/");
                String[] meses = {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"};
                String dia = numeroEnLetras(Integer.parseInt(dividirFecha[0]));
                String mes = meses[Integer.parseInt(dividirFecha[1]) - 1];

                fechaletra = dividirFecha[0] + " de " + mes + " de " + dividirFecha[2];
                if (opc == 1) {
                    fechaletra = dia + " (" + dividirFecha[0] + ") d??as del mes de " + mes + " de " + dividirFecha[2];
                }
            } catch (Exception ex) {
                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                throw new Exception(" " + ex.getMessage());
            }

        }
        return fechaletra;
    }

    private String getSoportesColciencias(List<Map<String, String>> listadocentexArticuloColciencias, int k) {

        String soportes = Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("NOMBRE_SOLICITUD"))
                + "; de la " + Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("REVISTA/EVENTO/EDITORIAL")) + "; ISSN:"
                + listadocentexArticuloColciencias.get(k).get("ISSN/ISBN") + "; " + Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("FECHA_PUBLICACION/REALIZACION")) + "; (" + Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("PUBLINDEX"))
                + "); " + listadocentexArticuloColciencias.get(k).get("No._AUTORES") + "autor(es). \n";
        return soportes;

    }

    private String getDecisionColciencias(List<Map<String, String>> listadocentexArticuloColciencias) throws Exception {
        String respuesta = "";
        String respuestaxEstado = "";
        double sumatoria_puntos = 0;
        for (int f = 0; f < listadocentexArticuloColciencias.size(); f++) {
            try {
                sumatoria_puntos += Double.parseDouble(ValidarNumero(listadocentexArticuloColciencias.get(f).get("PUNTOS").replace(",", ".")));
            } catch (Exception ex) {
                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                throw new Exception(" " + ex.getMessage());
            }

            respuestaxEstado += (listadocentexArticuloColciencias.size() > 1 ? " por el " + (getNombreNumero((f + 1), "el") + " art??culo") : "");
            if (listadocentexArticuloColciencias.get(f).get("RESPUESTA_CIARP").equals("Aprobado")) {

                respuestaxEstado += " asignarle "
                        + (listadocentexArticuloColciencias.get(f).get("SEXO").toUpperCase().equals("M")
                        ? "al" : "a la") + " "
                        + "docente "
                        + getNumeroDecimal(listadocentexArticuloColciencias.get(f).get("PUNTOS"))
                        + " (" + listadocentexArticuloColciencias.get(f).get("PUNTOS") + ") " + listadocentexArticuloColciencias.get(f).get("TIPO_PUNTAJE");
                if (listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD").length() > 10) {
                    respuestaxEstado += listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD") + ".";
                } else {
                    respuestaxEstado += " a partir de " + fechaEnletras(listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD"), 0) + ".";
                }

                if (!Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("NORMA")).equals("#N/A")) {
                    respuestaxEstado += " considerando que el art??culo"
                            + " corresponde a un(a) " + listadocentexArticuloColciencias.get(f).get("SUBTIPO_PRODUCTO")
                            + "(" + listadocentexArticuloColciencias.get(f).get("NORMA") + ")";
                }
                if (Integer.parseInt(ValidarNumero(listadocentexArticuloColciencias.get(f).get("No._AUTORES"))) > 3) {
                    if (Integer.parseInt(ValidarNumero(listadocentexArticuloColciencias.get(f).get("No._AUTORES"))) < 6) {
                        respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal b; numeral III, art??culo 10 del Decreto 1279 de 2002).";
                    } else {
                        respuestaxEstado += " y teniendo en cuenta el n??mero de autores (literal c; numeral III, art??culo 10 del Decreto 1279 de 2002).";
                    }
                }

            } else if (listadocentexArticuloColciencias.get(f).get("RESPUESTA_CIARP").equals("Rechazado")) {
                respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("DECISION"));
            }

        }

        String add = (listadocentexArticuloColciencias.size() > 1
                ? " Para un total de " + sumatoria_puntos + " salariales por la productividad presentada."
                : "");

        respuesta = respuestaxEstado + add;

        return respuesta;
    }

    private String getArticuloxTipoProducto(String tipo) {
        String ret = "el";

        for (Map<String, String> obj : jerarquiaProducto) {
            if (obj.get("PRODUCTO").equals(tipo)) {
                ret = obj.get("ARTICULO");
                break;
            }
        }

        return ret;
    }

    private String getDatoJerarquiaProducto(String datoB, String KeyB, String KeyR) {
        String ret = "";

        for (Map<String, String> obj : jerarquiaProducto) {
            if (obj.get(KeyB).equals(datoB)) {
                ret = obj.get(KeyR);
                break;
            }
        }

        return ret;
    }

    private String ComillasSoporte(String datosSoporte) {
        String ret = datosSoporte;
        ret = ret.replace("\"\"", "<::>");
        ret = ret.replace("\"", "");
        ret = ret.replace("<::>", "\"");
        return ret;
    }

    private String getNombreNumero(int numero, String articulo) {
        String nombre = numeroOrdinales(numero);

        if ("LA".equals(articulo.toUpperCase())) {
            nombre = nombre.substring(0, nombre.length() - 1) + "a";
        } else {
            if (numero == 1 || numero == 3) {
                nombre = nombre.substring(0, nombre.length() - 1);
            }
        }

        return nombre;
    }

    private String FormatoCedula(String cedula) {
        if (cedula.indexOf(",") >= 0) {
            String[] Cedul = cedula.split(",");
            cedula = Cedul[0];
        }
        String ret = "";
        int suma = 0;
        for (int i = cedula.length() - 1; i >= 0; i--) {
            if (suma == 3) {
                suma = 0;
                ret = "." + ret;
            }
            suma++;
            ret = "" + cedula.charAt(i) + ret;
        }

        return ret;
    }

    public String ValidarNumeroDec(String valor) {
        String retorno = "";

        if (valor.indexOf(",") > -1) {

            valor = valor.replace(",", ".");
        } else {
            retorno = valor;
        }

        if (valor.indexOf(".") > -1) {
            Double dat = Double.parseDouble(valor);
            DecimalFormat df = new DecimalFormat("0.0");
            valor = df.format(dat);
            valor = valor.replace(".", ",");
            String[] daot = valor.split(",");

            if (daot[1].equals("0")) {
                retorno = daot[0];
            } else {
                retorno = valor;
            }

        }

        return retorno;

    }

}
