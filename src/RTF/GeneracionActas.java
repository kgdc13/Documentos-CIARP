/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package RTF;

//import Configuracion.SystemParam;
import Excel.ControlArchivoExcel;
import Excel.gestorInformes;
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
 * @author rjulio
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
        auxiliarJerarquia.put("NORMA", "Artículo 14 del Acuerdo Superior N° 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ascenso_en_el_Escalafon_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ascenso en el Escalafón Docente");
        auxiliarJerarquia.put("NORMA", "Artículo 27 del Acuerdo Superior N° 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Titulacion");
        auxiliarJerarquia.put("NPRODUCTO", "Titulación");
        auxiliarJerarquia.put("NORMA", "Artículo 7 del Decreto 1279 del 2002, Artículo Primero del Acuerdo 001 de 2004 del Grupo de Seguimiento al Decreto 1279 de 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Articulo");
        auxiliarJerarquia.put("NPRODUCTO", "Artículo");
        auxiliarJerarquia.put("NORMA", "Literal a. numeral I, Artículo 10 y literal a. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Video_Cinematograficas_o_Fonograficas");
        auxiliarJerarquia.put("NPRODUCTO", "Producción de Video Cinematográfica o Fonográfica");
        auxiliarJerarquia.put("NORMA", "Literal b. numeral I, Artículo 10; literal b. numeral I, Artículo 24; literal a. numeral I; literal a. numeral II, Articulo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Libro");
        auxiliarJerarquia.put("NORMA", "Literales c, d, e, Artículo 10 y literales c, d, e. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Capitulo_de_Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Capítulo de Libro");
        auxiliarJerarquia.put("NORMA", "Literales c, d, e, Artículo 10 y literales c, d, e. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Premios_Nacionales_e_Internacionales");
        auxiliarJerarquia.put("NPRODUCTO", "Premio Nacional o Internacional");
        auxiliarJerarquia.put("NORMA", "Literal f. numeral I, Artículo 10 y literal g. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Patente");
        auxiliarJerarquia.put("NPRODUCTO", "Patente");
        auxiliarJerarquia.put("NORMA", "Literal g. numeral I, Artículo 10 y literal h. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traduccion_de_Libros");
        auxiliarJerarquia.put("NPRODUCTO", "Traducción de Libro");
        auxiliarJerarquia.put("NORMA", "Literal h. numeral I, Artículo 10 y literal f. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Obra_Artistica");
        auxiliarJerarquia.put("NPRODUCTO", "Obra Artística");
        auxiliarJerarquia.put("NORMA", "Literal i. numeral I, Artículo 10; literal i. numeral I, Artículo 24; literal b. numeral I; literal g. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_Tecnica");
        auxiliarJerarquia.put("NPRODUCTO", "Producción Técnica");
        auxiliarJerarquia.put("NORMA", "Literal j. numeral I, Artículo 10 y literal j. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Software");
        auxiliarJerarquia.put("NPRODUCTO", "Producción de Software");
        auxiliarJerarquia.put("NORMA", "Literal k. numeral I, Artículo 10 y literal k. numeral I, Artículo 24 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ponencias_en_Eventos_Especializados");
        auxiliarJerarquia.put("NPRODUCTO", "Ponencia en Evento Especializado");
        auxiliarJerarquia.put("NORMA", "Literal c. numeral I, literal b. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Publicaciones_Impresas_Universitarias");
        auxiliarJerarquia.put("NPRODUCTO", "Publicación Impresa Universitaria");
        auxiliarJerarquia.put("NORMA", "Literal d. numeral I, literal c. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Estudios_Posdoctorales");
        auxiliarJerarquia.put("NPRODUCTO", "Estudio Posdoctoral");
        auxiliarJerarquia.put("NORMA", "Literal e. numeral I, literal d. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Reseñas_Críticas");
        auxiliarJerarquia.put("NPRODUCTO", "Reseña Crítica");
        auxiliarJerarquia.put("NORMA", "Literal f. numeral I, literal e. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traducciones");
        auxiliarJerarquia.put("NPRODUCTO", "Traducción");
        auxiliarJerarquia.put("NORMA", "Literal g. numeral I, literal f. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Direccion_de_Tesis");
        auxiliarJerarquia.put("NPRODUCTO", "Dirección de Tesis");
        auxiliarJerarquia.put("NORMA", "Literal h. numeral I, literal h. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Evaluacion_como_par");
        auxiliarJerarquia.put("NPRODUCTO", "Evaluación como par");
        auxiliarJerarquia.put("NORMA", "Literal i. numeral I, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        jerarquiaProducto.add(auxiliarJerarquia);
    }

    public String Encode() {
        String cifrado = "" + System.currentTimeMillis();
        return cifrado;
    }

    public Map<String, String> GenerarActas(String ruta) throws DocumentException, IOException {
        System.out.println("¨*****************GenerarActas****************" + ruta);
        this.ruta = ruta;
        ControlArchivoExcel con = new ControlArchivoExcel();
        List<Map<String, String>> listaDatos = new ArrayList<>();
        
        //<editor-fold defaultstate="collapsed" desc="Lectura Orden del día">
        String extP = ruta.substring(ruta.lastIndexOf(".") + 1);
        System.out.println("************************EMPIEZA LECTURA ORDEN DEL DIA");
        System.out.println("rutaPuntos--->"+ruta);
        System.out.println("extP--->"+extP);
        if (extP.equals("xlsx")) {
            listaDatos = con.LeerExcelDesdeAct(ruta, 2, "ORDEN DEL DIA");
            listaDatosacta = con.LeerExcelParametrosAct(ruta, 1, 1,"ORDEN DEL DIA");
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
        
//        String[] keys = new String[]{};
//        int nkey = 0;
//        if (br != null) {
//            try {
//                while ((lineaDeTexto = br.readLine()) != null) {
//                    if (nkey == 0) {
//                        String d[] = lineaDeTexto.split("\\t");
//                        listaDatosacta.put("No_ACTA", d[1]);
//                        listaDatosacta.put("FECHA_ACTA", d[3]);
//                        listaDatosacta.put("HORA_INICIO", d[5]);
//                        listaDatosacta.put("HORA_FIN", d[7]);
//                        listaDatosacta.put("ACTA_ANTERIOR", d[9]);
//                        listaDatosacta.put("FECHA_ACTA_ANTERIOR", d[11]);
//                        nkey = 1;
//                    } else if (nkey == 1) {
//                        keys = lineaDeTexto.split("\\t");
//                        nkey = 2;
//                    } else {
//                        String d[] = lineaDeTexto.split("\\t");
//                        if(d.length>0){
//                        for (int k = 0; k < d.length; k++) {
//                            System.out.println(" K = " + k + "Valor " + d[k]);
//                        }
//                        Map<String, String> mapeado = new HashMap<>();
//                        System.out.println(" Size KEYS= " + keys.length + "Size D=  " + d.length);
//                        for (int i = 0; i < keys.length; i++) {
//                            mapeado.put(keys[i], d[i]);
//                        }
//                        listaDatos.add(mapeado);
//                        }
//                    }
//
//                    //String d[] = lineaDeTexto.split("\\t");
//                    //GENERACION DOCUMENTO
////                    
////                    
//                    //</editor-fold>
//                }
                URL = "C:\\CIARP\\acta_" + listaDatosacta.get("No_ACTA") + ".rtf";
                GeneracionDocumentoActa(listaDatos, documento);
                for (int j = 0; j < listaDatos.size(); j++) {
                    // System.out.println("***********************inicio j: "+j+"***************************************");

                    for (Map.Entry<String, String> entry : listaDatos.get(j).entrySet()) {
                        String key = entry.getKey();
                        String value = entry.getValue();
                        //  System.out.println("KEY= "+key+", valor= "+value);

                    }
                    //System.out.println("*************************finalizacion J ="+j+"*************************************");
                }

                documento.close();
                int result = JOptionPane.showConfirmDialog(null, "¿Desea abrir el documento?");
                if (result == JOptionPane.YES_OPTION) {
                    Desktop.getDesktop().open(new File(URL));
                }
                gestorInformes gi = new gestorInformes(Datosacta, DatosNumeralesActa);
                gi.iniciar();
                
                System.out.println("documento.getPageNumber()--->" + documento);
                return respuestaerr;
            } 


    private Map<String, String> GeneracionDocumentoActa(List<Map<String, String>> listaDatos, Document documento) throws FileNotFoundException, BadElementException, IOException, DocumentException {
        BaseFont calibriFont = BaseFont.createFont("C:\\windows\\Fonts\\CALIBRI.TTF", "Cp1252", true);
        BaseFont arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
        RtfWriter2.getInstance(documento, new FileOutputStream(URL));
        

        documento.open();
        List<Map<String, String>> listaCorrespondencia = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Revision_de_la_correspondencia"});
        List<Map<String, String>> listaProposiciones = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Proposiciones_y_varios"});
        List<Map<String, String>> listaColciencias = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + "Art_Col"});
        //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">
        Font fh1 = new Font(arialFont);
        fh1.setSize(11);
        fh1.setStyle("bold");
        fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("bold");

        Font fh3 = new Font(arialFont);
        fh3.setSize(11);
        fh3.setColor(Color.BLACK);

        Font fh4 = new Font(arialFont);
        fh4.setSize(9);
        fh4.setColor(Color.BLACK);
        //fh2.setStyle("bold");
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
            celda = new Cell(new Paragraph("COMITÉ INTERNO DE ASIGNACIÓN Y RECONOCIMIENTO DE PUNTAJE\n", fh1));
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
            celda = new Cell(new Paragraph("Acta N° " + listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0), fh4));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", ""+ex.getMessage());
            respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha del acta de la primera fila del documento");
            respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
        }
            celda.setBorder(0);
            celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
            celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
            celda.setBorderWidthTop(2);
    //        celda.setColspan(1);
            footertable.addCell(celda);
            celda = new Cell(new RtfPageNumber(fh4));
            celda.setBorder(0);
            celda.setVerticalAlignment(Cell.ALIGN_TOP);
            celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
            celda.setBorderWidthTop(2);
            footertable.addCell(celda);

            celda = new Cell(new Paragraph("/", fh4));
            celda.setBorder(0);
            celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
            celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
            celda.setBorderWidthTop(2);
            footertable.addCell(celda);

            celda = new Cell(new RtfTotalPageNumber(fh4));
            celda.setBorder(0);
            celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
            celda.setBorderWidthTop(2);
            footertable.addCell(celda);
        //</editor-fold>
        

        RtfHeaderFooter header = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);


        documento.setHeader(header);
        documento.setFooter(footer);
        //documento.setFooter(footer);
        //</editor-fold>

//            fh1 = new Font(fh2);
//            fh1.setSize(11);
//            fh1.setStyle(Font.NORMAL);
//            fh1.setColor(Color.black);
//
//            fh2.setSize(11);
//            fh2.setStyle(Font.BOLD);
//            fh2.setColor(Color.black);
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        //<editor-fold defaultstate="collapsed" desc="CONTENIDO DEL DOCUMENTO">
        //<editor-fold defaultstate="collapsed" desc="CUERPO ACTA">
        p.setAlignment(centrado);
        p.setFont(fh2);
        
        String[] año = listaDatosacta.get("FECHA_ACTA").split("/");
       p.add("\nACTA " + listaDatosacta.get("No_ACTA") + " DE "+ año[2]);
        
        //p.add("\nACTA " + listaDatosacta.get("No ACTA") + " DE 2019 ");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(fh1);
        p.setAlignment(centrado);
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(fh3);
        p.setAlignment(justificado);
        try {
            p.add("En Santa Marta, a los " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 1) + ", se reunieron en sesión ordinaria los miembros del ");
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", ""+ex.getMessage());
            respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha del acta de la primera fila del documento");
            respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
                        
        }
        Chunk c = new Chunk("COMITÉ INTERNO DE ASIGNACIÓN Y RECONOCIMIENTO DE PUNTAJE ", fh2);
        p.add(c);
        String orden = "convocados por el Presidente de éste órgano colegiado, para tratar el siguiente orden del día:\n"
                + "\n"
                + "1. Verificación del Quórum.\n"
                + "\n"
                + "2. Aprobación del orden del día.\n"
                + "\n"
                + "3. Lectura y aprobación del acta anterior.\n"
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
        c = new Chunk("DESARROLLO\n", fh2);
        p.add(c);

        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(fh3);
        p.add("\n Se da inicio a la sesión, en el despacho del Vicerrector Académico, siendo las " + listaDatosacta.get("HORA_INICIO") + "\n \n");
        c = new Chunk("1. VERIFICACIÓN DEL QUÓRUM ", fh2);
        p.add(c);
        p.add("se verifica el quórum para deliberar contando con los siguientes asistentes:\n"
                + "\n"
                + "1. Oscar Humberto Garcia Vargas - Presidente del Comité\n"
                + "2. Jorge Enrique Elias Caro – Vicerrector de Investigación\n"
                + "3. Haidy Rocio Oviedo Cordoba (Representante de los docentes del Área Ciencias de la Salud)\n"
                + "4. Rocio Del Pilar Garcia Urueña (Representante de los docentes del Área de Matemáticas y Ciencias Naturales)\n"
                + "5. Rolando Enrique Escorcia Caballero (Representante de los docentes del Área de Ciencias de la Educación)\n"
                + "6. Marco Francisco Gaviria Rueda (Representante de los docentes del Área de Arquitectura, Bellas Artes y Afines)\n"
                + "7. Karen Gishelle Buelvas Ferreira (Profesional Especializado de la Vicerrectoría Académica)\n"
                + "8. Kennys De Los Reyes Castillo (Profesional Universitario de la Vicerrectoría Académica)\n"
                + "9. Diana Paola Orozco Tete (Contratista)\n"
                + "\n"
                + "Considerando lo establecido en el artículo 5to del Acuerdo Superior N° 021 de 2009, que dispone la composición y requisitos de los miembros del Comité Interno de Asignación y Reconocimiento de Puntaje atendiendo los principios de eficacia, economía y celeridad que enmarcan todas las actuaciones administrativas, éste órgano colegiado, decide por unanimidad designar a la Profesional Especializado de la Vicerrectoría Académica Karen Gishelle Buelvas Ferreira, como secretaria técnica para que se encargue de la parte operativa del CIARP.  \n"
                + "\n");
        c = new Chunk("2. APROBACIÓN DEL ORDEN DEL DÍA ", fh2);
        p.add(c);
        p.add("se aprueba el orden del día.\n"
                + "\n");
        c = new Chunk("3. LECTURA Y APROBACIÓN DEL ACTA ANTERIOR ", fh2);
        p.add(c);
        try {
            p.add("leída el Acta N° " + listaDatosacta.get("ACTA_ANTERIOR") + " de fecha " + fechaEnletras(listaDatosacta.get("FECHA_ACTA_ANTERIOR"), 0) + " es aprobada por unanimidad por los miembros del Comité.  \n"
                    + "\n");
        } catch (Exception ex) {
            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
            respuestaerr.put("ESTADO", "ERROR");
            respuestaerr.put("MENSAJE", ""+ex.getMessage());
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
            c = new Chunk(indxgrado1 + ". ESTUDIO DE LA CORRESPONDENCIA: \n", fh2);
            p.add(c);
            documento.add(p);
            indxgrado2 = 1;
            for (int l = 0; l < listaCorrespondencia.size(); l++) {

                p = new Paragraph(10);
                p.setFont(fh3);
                p.setAlignment(justificado);
                String numeral = indxgrado1 + "." + (indxgrado2++);
                
                //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                    datos1 = new HashMap<>();
                    datos1.put("TIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("TIPO_PRODUCTO")));
                    datos1.put("SUBTIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("SUBTIPO_PRODUCTO")));
                    datos1.put("IDDOCENTE", ""+listaCorrespondencia.get(l).get("No._IDENTIFICACION"));
                    datos1.put("DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_DOCENTE")));
                    datos1.put("NOMBRE_SOLICITUD", ""+Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_SOLICITUD")));
                    datos1.put("RESPUESTA_CIARP", ""+Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("RESPUESTA_CIARP")));
                    datos1.put("NUMERAL", ""+numeral);
                    datos1.put("POS", "1");
                    datos2.add(datos1);
                //</editor-fold>
                    
                c = new Chunk(numeral+". ", fh2);
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("NOMBRE_SOLICITUD"))) + " \n");
                c = new Chunk("Decisión: ", fh2);
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaCorrespondencia.get(l).get("DECISION"))) + ".\n");
                documento.add(p);
            }

        }

        indxgrado1++;

        p = new Paragraph(10);
        c = new Chunk(indxgrado1 + ". ESTUDIO DE LAS SOLICITUDES DE DOCENTES: \n", fh2);
        p.add(c);
        documento.add(p);
        indxgrado2 = 0;
        for (int i = 0; i < jerarquiaProducto.size(); i++) {
            List<Map<String, String>> listadatosxTipoProducto = data_list(3, listaDatos, new String[]{"TIPO_PRODUCTO<->" + jerarquiaProducto.get(i).get("PRODUCTO")});
            System.out.println("JERARQUIA " + jerarquiaProducto.get(i).get("PRODUCTO") + "  EN LA POICION " + i + "tiene " + listadatosxTipoProducto.size() + " Datos");

            if (listadatosxTipoProducto.size() > 0) {
                Datosacta.put("NOMBRE_ARCHIVO", "numerales_acta_");
                Datosacta.put("NUMACTA", listadatosxTipoProducto.get(0).get("ACTA"));
                indxgrado2++;
                p = new Paragraph(10);
                p.setFont(fh2);
                p.setAlignment(justificado);
                p.add(indxgrado1 + "." + indxgrado2 + ". Estudio de solicitudes por " + jerarquiaProducto.get(i).get("PRODUCTO").replaceAll("_", " ") + ":\n");
                documento.add(p);

                List<Map<String, String>> listadocentexTipoProducto = data_list(1, listadatosxTipoProducto, new String[]{"No._IDENTIFICACION"});
                indxgrado3 = 0;
                for (int j = 0; j < listadocentexTipoProducto.size(); j++) {
                    List<Map<String, String>> listadatosdocentexTipoProducto = data_list(3, listadatosxTipoProducto, new String[]{"No._IDENTIFICACION<->" + listadocentexTipoProducto.get(j).get("No._IDENTIFICACION")});
                    indxgrado3 += 1;
                    p = new Paragraph(10);
                    p.setFont(fh3);
                    p.setAlignment(justificado);
                    String numeral = indxgrado1 + "." + indxgrado2 + "." + indxgrado3 ;
                    c = new Chunk(numeral+". " + Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOMBRE_DOCENTE")) + ":\n", fh2);
                    p.add(c);
                    c = new Chunk("Identificación: ", fh2);
                    p.add(c);
                    p.add(FormatoCedula(listadocentexTipoProducto.get(j).get("No._IDENTIFICACION")) + "\n" + Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("FACULTAD")) + ".\n");
                    c = new Chunk("Categoría: ", fh2);
                    p.add(c);
                    try {
                        p.add(listadocentexTipoProducto.get(j).get("CATEGORIA_DOCENTE") + " - " + listadocentexTipoProducto.get(j).get("TIPO_VINCULACION") + " desde el " + fechaEnletras(listadocentexTipoProducto.get(j).get("FECHA_INGRESO"), 0) + ".\n");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                        respuestaerr.put("ESTADO", "ERROR");
                        respuestaerr.put("MENSAJE", ""+ex.getMessage());
                        respuestaerr.put("LINEA_ERROR_FECHA", "En la fecha de vinculación del docente " +listadocentexTipoProducto.get(j).get("NOMBRE_DOCENTE"));
                        respuestaerr.put("DEFINICION_FORMATO", "El formato de fecha debe ser dd/mm/aaaa");
                    }
                    c = new Chunk("Tipo de solicitud: ", fh2);
                    p.add(c);
                    System.out.println(" PROBANDO A VER QUE PASA AQUI  "+listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO"));
                    p.add((listadatosxTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")||listadatosxTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Ingreso_a_la_Carrera_Docente")?
                            "Puntos por":"Puntos por la publicación de " + listadatosdocentexTipoProducto.size()) + " " +
                            (listadatosdocentexTipoProducto.size()>1?
                                    jerarquiaProducto.get(i).get("PRODUCTO").replaceAll("_", " "):
                                    jerarquiaProducto.get(i).get("NPRODUCTO"))
                            + ". (" + jerarquiaProducto.get(i).get("NORMA") + ").\n");
                    //documento.add(p);

                    for (int k = 0; k < listadatosdocentexTipoProducto.size(); k++) {
                        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                            datos1 = new HashMap<>();
                            datos1.put("TIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO")));
                            datos1.put("SUBTIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SUBTIPO_PRODUCTO")));
                            datos1.put("IDDOCENTE", ""+listadatosdocentexTipoProducto.get(k).get("No._IDENTIFICACION"));
                            datos1.put("DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_DOCENTE")));
                            datos1.put("NOMBRE_SOLICITUD", ""+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD")));
                            datos1.put("RESPUESTA_CIARP", ""+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("RESPUESTA_CIARP")));
                            datos1.put("NUMERAL", ""+numeral);
                            datos1.put("POS", ""+(k+1));
                            datos2.add(datos1);
                        //</editor-fold>
                        //p= new Paragraph();
                        p.setFont(fh3);
                        p.setAlignment(justificado);
                        c = new Chunk("Soporte" + (listadatosdocentexTipoProducto.size() > 1 ? 
                                (" " + getNombreNumero(k + 1, jerarquiaProducto.get(i).get("ARTICULO")) + " " + jerarquiaProducto.get(i).get("NPRODUCTO") + ": ") : 
                                ": ")
                                , fh2);
                        p.add(c);
                        String soportes = getSoportes(listadatosdocentexTipoProducto, k);
                        p.add(soportes + "\n");
                        //documento.add(p);
                    }

                    String decision;
                    try {
                        decision = getDecision(listadatosdocentexTipoProducto);
                         p.setFont(fh3);
                    p.setAlignment(justificado);
                    c = new Chunk("Decisión: ", fh2);
                    p.add(c);
                    p.add("Revisada la documentación y después de analizar lo establecido en las normas, el Comité decide" + decision + "\n");
                    //documento.add(p);
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                        respuestaerr.put("ESTADO", "ERROR");
                        respuestaerr.put("MENSAJE", ""+ex.getMessage());
                        System.out.println(" ERROROOO "+ex.getMessage());
                        respuestaerr.put("LINEA_ERROR_DOCENTE", ""+listadatosdocentexTipoProducto.get(0).get("NOMBRE_DOCENTE"));
                        respuestaerr.put("LINEA_ERROR_PRODUCTO", ""+listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO"));
                         if(listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD").length()<=100){
                        respuestaerr.put("NOMBRE_PRODUCTO", ""+listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD"));
                        }else{
                        respuestaerr.put("NOMBRE_PRODUCTO", ""+listadatosdocentexTipoProducto.get(0).get("NOMBRE_SOLICITUD").substring(0, 100));
                        }
                        return respuestaerr;
                    }
                    //p= new Paragraph();
                   

                    if (!Utilidades.Utilidades.decodificarElemento(listadocentexTipoProducto.get(j).get("NOTA")).equals("N/A")) {
                        p.setFont(fh3);
                        p.setAlignment(justificado);
                        c = new Chunk("Nota: ", fh2);
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
            c = new Chunk(indxgrado1 + ". REVISIÓN Y ASIGNACIÓN DE PUNTOS A LOS DOCENTES CUYOS ARTÍCULOS QUEDARON PENDIENTES POR HOMOLOGACIÓN DE COLCIENCIAS PARA EL AÑO 2019: \n", fh2);
            p.add(c);
            documento.add(p);
            indxgrado2 = 1;
            p= new Paragraph(10);
            p.add("Teniendo en cuenta la respuesta brindada por el Grupo de Seguimiento al Régimen Salarial y Prestacional de los Docentes de universidades Públicas a la consulta: "
                    + "¿Se pueden reconocer puntos salariales que sean solicitados y estudiados antes de la actualización de la base de datos de Colciencias"
                    + " y hacer el pago de estos una vez se conozca la categoría en que se encuentra la revista; y de igual manera pagarse los meses transcurridos"
                    + " a partir del estudio de la solicitud?\n");
            p=new Paragraph(10);
            c = new Chunk("RESPUESTA:",fh2);
            p.add(c);
            p= new Paragraph(10);
            p.add("“En este caso se reconoce provisionalmente el puntaje hasta tanto se publique por Colciencias por la clasificación de las revistas, se aclara que no hay retroactividad,"
                    + " y solo se asigna el puntaje hasta tanto que se encuentre en firme la clasificación de las revistas de cada institución, se deberá tener en cuenta que la Homologación "
                    + "se surte en el último trimestre de cada año.\n" +
                      "Así las cosas se hará efectivo el reconocimiento solo a partir del momento en el que se haga público el listado"
                    + " de revistas extranjeras homologadas para la respectiva vigencia” \n" +
                      " \n" +
                    "Y que el Comité se acogió a esta respuesta; se hace asignación de los puntos salariales que habían quedado pendientes "
                    + "a los docentes que solicitaron puntos por los artículos publicados en las revistas que no se encontraban en la actualización de"
                    + " base de datos de Colciencias, según la siguiente relación: \n");
            
            System.out.println("listaColciencias.size()-->"+listaColciencias.size());
            System.out.println("listaColciencias.size()-->"+listadocentexArticuloColciencias.size());
            
            for(int l = 0; l < listadocentexArticuloColciencias.size(); l++){
                List<Map<String, String>> listainfoDocentesColciencias = data_list(3, listaColciencias, new String[]{"No._IDENTIFICACION<->" + listadocentexArticuloColciencias.get(l).get("No._IDENTIFICACION")});
                p = new Paragraph(10);
                p.setFont(fh3);
                p.setAlignment(justificado);
                String numeral = indxgrado1 + "." + (indxgrado2++);
                c=new Chunk(numeral + ". DOCENTE: " + listadocentexArticuloColciencias.get(l).get("NOMBRE_DOCENTE")+"\n",fh2);
                p.add(c);
                
                for (int k = 0; k < listainfoDocentesColciencias.size(); k++) {
                    //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                            datos1 = new HashMap<>();
                            datos1.put("TIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("TIPO_PRODUCTO")));
                            datos1.put("SUBTIPO_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("SUBTIPO_PRODUCTO")));
                            datos1.put("IDDOCENTE", ""+listainfoDocentesColciencias.get(k).get("No._IDENTIFICACION"));
                            datos1.put("DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("NOMBRE_DOCENTE")));
                            datos1.put("NOMBRE_SOLICITUD", ""+Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("NOMBRE_SOLICITUD")));
                            datos1.put("RESPUESTA_CIARP", ""+Utilidades.Utilidades.decodificarElemento(listainfoDocentesColciencias.get(k).get("RESPUESTA_CIARP")));
                            datos1.put("NUMERAL", ""+numeral);
                            datos1.put("POS", ""+(k+1));
                            datos2.add(datos1);
                        //</editor-fold>
                    //p= new Paragraph();
                    p.setFont(fh3);
                    p.setAlignment(justificado);
                    c = new Chunk("Artículo "+(listainfoDocentesColciencias.size() > 1 ? "N° "+(k + 1)+": " : ": "), fh2);
                    p.add(c);
                    String soportes = getSoportesColciencias(listainfoDocentesColciencias, k);
                    p.add(soportes + "");
                    //documento.add(p);
                }
                                
                c = new Chunk("DECISIÓN: ", fh2);
                p.add(c);
                
                try {
                    String decision;
                    decision = getDecisionColciencias(listainfoDocentesColciencias);
                    p.add(decision+"\n");
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                    respuestaerr.put("ESTADO", "ERROR");
                    respuestaerr.put("MENSAJE", ""+ex.getMessage());
                    respuestaerr.put("LINEA_ERROR_DOCENTE", ""+listainfoDocentesColciencias.get(l).get("NOMBRE_DOCENTE"));
                    respuestaerr.put("LINEA_ERROR_PRODUCTO", ""+listainfoDocentesColciencias.get(l).get("TIPO_PRODUCTO"));
                    if(listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD").length()<=100){
                        respuestaerr.put("NOMBRE_PRODUCTO", ""+listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD"));
                        }else{
                        respuestaerr.put("NOMBRE_PRODUCTO", ""+listainfoDocentesColciencias.get(l).get("NOMBRE_SOLICITUD").substring(0, 100));
                        }
                    return respuestaerr;
                }
                
                documento.add(p);
                
                if (!Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(l).get("NOTA")).equals("N/A")) {
                    p.setFont(fh3);
                    p.setAlignment(justificado);
                    c = new Chunk("Nota: ", fh2);
                    p.add(c);
                    p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(l).get("NOTA"))) + "\n");

                }
            }

        }

        if (listaProposiciones.size() > 0) {
            indxgrado1++;
            indxgrado2 = 0;
            p = new Paragraph(10);
            p.setFont(fh3);
            p.setAlignment(justificado);
            c = new Chunk(indxgrado1 + ". PROPOSICIONES Y VARIOS: \n\n", fh2);
            p.add(c);
            for (int y = 0; y < listaProposiciones.size(); y++) {
                indxgrado2++;
                p.add(indxgrado1 + "." + indxgrado2 + " " + ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaProposiciones.get(y).get("NOMBRE_SOLICITUD"))) + "\n");
                c = new Chunk("Decisión: ", fh2);
                p.add(c);
                p.add(ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listaProposiciones.get(y).get("DECISION"))) + "\n");
            }
            
            documento.add(p);
        }
        DatosNumeralesActa.add(datos2);
        indxgrado1++;
        p = new Paragraph(10);
        p.setFont(fh3);
        p.setAlignment(justificado);
        c = new Chunk(indxgrado1 + ". CIERRE: \n\n", fh2);
        p.add(c);
        p.add(" Siendo las " + listaDatosacta.get("HORA_FIN") + " se da por terminada la sesión");
        documento.add(p);
        respuestaerr.put("ESTADO", "OK");
        respuestaerr.put("MENSAJE", "Las cartas se generaron satisfactoriamente.");
//
return respuestaerr;
    }

//<editor-fold defaultstate="collapsed" desc="METODO GeneracionDocumento ANTERIOR">
    //    private static void GeneracionDocumento(List<Map<String, String>> listaDatos, Document documento) throws FileNotFoundException, BadElementException, IOException, DocumentException {
//        BaseFont calibriFont = BaseFont.createFont("C:\\windows\\Fonts\\CALIBRI.TTF", "Cp1252", true);
//        BaseFont arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
//        RtfWriter2.getInstance(documento, new FileOutputStream(URL));
//        
//        documento.open();
//        List<Map<String, String>> listaCorrespondencia = data_list(3, listaDatos, new String[]{"TIPO PRODUCTO<->" + "Revision_de_la_correspondencia"});
//        List<Map<String, String>> listaProposiciones = data_list(3, listaDatos, new String[]{"TIPO PRODUCTO<->" + "Proposiciones_y_varios"});
//       editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">
//        Font fh1 = new Font(arialFont);
//        fh1.setSize(11);
//        fh1.setStyle("bold");
//        fh1.setStyle("underlined");
//
//        Font fh2 = new Font(arialFont);
//        fh2.setSize(11);
//        fh2.setColor(Color.BLACK);
//        fh2.setStyle("bold");
//
//        Font fh3 = new Font(arialFont);
//        fh3.setSize(11);
//        fh3.setColor(Color.BLACK);
//
//        Font fh4 = new Font(arialFont);
//        fh4.setSize(9);
//        fh4.setColor(Color.BLACK);
//        //fh2.setStyle("bold");
//        Paragraph p = new Paragraph();
//        int justificado = Paragraph.ALIGN_JUSTIFIED;
//        int centrado = Paragraph.ALIGN_CENTER;
//
//        //Image imgR = Image.getInstance("T:\\PASOACOACTIVOAPI\\logos\\logorightapinc.png");
//        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
//
//        Table headerTable = new Table(1, 2);
//        headerTable.setWidth(100);
//
//        Cell celda = new Cell(imgL);
//        celda.setBorder(0);
//        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
//        headerTable.addCell(celda);
//        celda = new Cell(new Paragraph("COMITE INTERNO DE ASIGNACION Y RECONOCIMIENTO DE PUNTAJE", fh1));
//        celda.setBorder(0);
//        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
//        headerTable.addCell(celda);
////            celda = new Cell(imgR);
////            celda.setBorder(0);
////            celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
////            headerTable.addCell(celda);
//
////            celda = new Cell();
////            celda.setBorder(0);
////            headerTable.addCell(celda);
////            celda = new Cell(new Paragraph("SECRETARIA DE HACIENDA", fh2));
////            celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
////            celda.setBorder(0);
////            headerTable.addCell(celda);
////            celda = new Cell();
////            celda.setBorder(0);
////            headerTable.addCell(celda);
//        Table footertable = new Table(2, 1);
//        footertable.setWidth(100);
//        
//        
////        footertable.addColumns(1);
//        celda = new Cell(new Paragraph("Acta N°" + listaDatosacta.get("No ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0), fh4));
//        celda.setBorder(0);
//        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
//        celda.setBorderWidthTop(2);
////        celda.setColspan(1);
//        footertable.addCell(celda);
//        
//        
//        celda = new Cell(new Paragraph( new RtfPageNumber()+"-2/13-"+new RtfTotalPageNumber(), fh4));
//        
//        celda.setBorder(0);
//        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
//        celda.setBorderWidthTop(2);
//        footertable.addCell(celda);
////            celda = new Cell(new Paragraph("Calle 14 No. 2-49  Conmutador: +57 (5) 4382777 Fax +57 (5) 4382993", fh1));
////            celda.setBorder(0);
////            celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
////            footertable.addCell(celda);
////            celda = new Cell(new Paragraph("www.santamarta-magdalena.gov.co", fh1));
////            celda.setBorder(0);
////            celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
////            footertable.addCell(celda);
//
//        RtfHeaderFooter header = new RtfHeaderFooter(headerTable);
//        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);
//        
//        Paragraph pp=new Paragraph();
//        pp.add( "Page  ");
//        pp.add(new RtfPageNumber());
//        pp.add( " of  ");
//        pp.add(new RtfTotalPageNumber());
//        pp.setAlignment(Element.ALIGN_CENTER);
//        
//        RtfHeaderFooter footer1 = new RtfHeaderFooter(pp);
//        //documento.setFooter(footer1);
//        
//        
//        
//        documento.setHeader(header);
//        documento.setFooter(footer1);
//        //documento.setFooter(footer);
//        ///editor-fold>
//
////            fh1 = new Font(fh2);
////            fh1.setSize(11);
////            fh1.setStyle(Font.NORMAL);
////            fh1.setColor(Color.black);
////
////            fh2.setSize(11);
////            fh2.setStyle(Font.BOLD);
////            fh2.setColor(Color.black);
//        
//        p = new Paragraph();
//        justificado = Paragraph.ALIGN_JUSTIFIED;
//        centrado = Paragraph.ALIGN_CENTER;
//
//            editor-fold defaultstate="collapsed" desc="CONTENIDO DEL DOCUMENTO">
//        ditor-fold defaultstate="collapsed" desc="CUERPO ACTA">
//        p.setAlignment(centrado);
//        p.setFont(fh2);
//        p.add("\nACTA " + listaDatosacta.get("No ACTA") + " DE 2019 ");
//        documento.add(p);
//
//        p = new Paragraph();
//        p.setFont(fh1);
//        p.setAlignment(centrado);
//        p.add("\n");
//        documento.add(p);
//
//        p = new Paragraph();
//        p.setFont(fh3);
//        p.setAlignment(justificado);
//        p.add("En Santa Marta, a los " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 1) + ", se reunieron en sesión ordinaria los miembros del ");
//        Chunk c = new Chunk("COMITÉ INTERNO DE ASIGNACIÓN Y RECONOCIMIENTO DE PUNTAJE ", fh2);
//        p.add(c);
//        String orden = "convocados por el Presidente de éste órgano colegiado, para tratar el siguiente orden del día:\n"
//                        + "\n"
//                        + "1. Verificación del Quórum.\n"
//                        + "\n"
//                        + "2. Aprobación del orden del día.\n"
//                        + "\n"
//                        + "3. Lectura y aprobación del acta anterior.\n"
//                        + "\n";
//        indxgrado1 = 4;
//        if(listaCorrespondencia.size()>0){
//            
//            orden+= (indxgrado1++)+". Estudio de la Correspondencia\n"
//                    + "\n";
//        }
//        
//        orden+=(indxgrado1++)+". Solicitudes de los Docentes de Planta\n"
//                + "\n";
//        if(listaProposiciones.size()>0){
//            
//            orden+= (indxgrado1++)+". Proposiciones y varios\n"
//                    + "\n";
//        }
//        
//        orden+=(indxgrado1++)+". Cierre. \n\n";
//        
//        p.add(orden);
//        documento.add(p);
//        
//        p = new Paragraph();
//        p.setAlignment(centrado);
//        c = new Chunk("DESARROLLO\n", fh2);
//        p.add(c);
//        
//        documento.add(p);
//        
//        p = new Paragraph();
//        p.setAlignment(justificado);
//        p.setFont(fh3);
//        p.add("\n Se da inicio a la sesión, en el despacho del Vicerrector Académico, siendo las " + listaDatosacta.get("HORA_INICIO") + "\n \n");
//        c = new Chunk("1. VERIFICACIÓN DEL QUÓRUM ", fh2);
//        p.add(c);
//        p.add("se verifica el quórum para deliberar contando con los siguientes asistentes:\n"
//                + "\n"
//                + "1. Jose Rafael Vasquez Polo - Presidente del Comité\n"
//                + "2. Ernesto Amaru Galvis Lista – Vicerrector de Investigación\n"
//                + "3. Haidy Rocio Oviedo Cordoba (Representante de los docentes del Área Ciencias de la Salud)\n"
//                + "4. Nelson Virgilio Piraneque Gambasica (Representante Suplente de los docentes del Área de Ingeniería)\n"
//                + "5. Rolando Escorcia caballero (Representante de los docentes del Área de Ciencias de la Educación)\n"
//                + "6. Kennys De Los Reyes Castillo (Técnico Administrativo de la Vicerrectoría Académica)\n"
//                + "7. Mibaflor Cantillo Baraza (Contratista)\n"
//                + "\n"
//                + "Considerando lo establecido en el artículo 5to del Acuerdo Superior N° 021 de 2009, que dispone la composición y requisitos de los miembros del Comité Interno de Asignación y Reconocimiento de Puntaje atendiendo los principios de eficacia, economía y celeridad que enmarcan todas las actuaciones administrativas, éste órgano colegiado, decide por unanimidad designar a la Técnico Administrativo de la Vicerrectoría Académica Kennys De Los Reyes Castillo, como secretaria técnica para que se encargue de la parte operativa del CIARP.  \n"
//                + "\n");
//        c = new Chunk("2. APROBACIÓN DEL ORDEN DEL DÍA ", fh2);
//        p.add(c);
//        p.add("se aprueba el orden del día.\n"
//                + "\n");
//        c = new Chunk("3. LECTURA Y APROBACIÓN DEL ACTA ANTERIOR ", fh2);
//        p.add(c);
//        p.add("leída el Acta N° "+listaDatosacta.get("ACTA_ANTERIOR")+" de fecha "+fechaEnletras(listaDatosacta.get("FECHA_ACTA_ANTERIOR"),0)+" es aprobada por unanimidad por los miembros del Comité.  \n"
//                + "\n");
//        indxgrado1 = 3;
//
//        documento.add(p);
//        
//        if (listaCorrespondencia.size() > 0) {
//            indxgrado1++;
//            p = new Paragraph();
//            c = new Chunk(indxgrado1 + ". ESTUDIO DE LA CORRESPONDENCIA: \n", fh2);
//            p.add(c);
//            documento.add(p);
//            indxgrado2 =1;
//            for (int l = 0; l < listaCorrespondencia.size(); l++) {
//                
//                p = new Paragraph();
//                p.setFont(fh3);
//                p.setAlignment(justificado);
//                p.add(indxgrado1 + "." + (indxgrado2++) + ". Se estudia la solicitud de " + listaCorrespondencia.get(l).get("NOMBRE DOCENTE") + " quien "
//                        + listaCorrespondencia.get(l).get("NOMBRE SOLICITUD") + ". \n");
//                c = new Chunk("Decisión: ", fh2);
//                p.add(c);
//                p.add(listaCorrespondencia.get(l).get("DECISION") + "\n");
//                documento.add(p);
//            }
//
//        }
//
//        indxgrado1++;
//        
//        p = new Paragraph();
//        c = new Chunk(indxgrado1 + ". ESTUDIO DE LAS SOLICITUDES DE DOCENTES DE PLANTA: \n", fh2);
//        p.add(c);
//        documento.add(p);
//        indxgrado2=0;
//        for (int i = 0; i < jerarquiaProducto.size(); i++) {
//            List<Map<String, String>> listadatosxTipoProducto = data_list(3, listaDatos, new String[]{"TIPO PRODUCTO<->" + jerarquiaProducto.get(i).get("PRODUCTO")});
//            System.out.println("JERARQUIA " + jerarquiaProducto.get(i).get("PRODUCTO") + "  EN LA POICION " + i + "tiene " + listadatosxTipoProducto.size() + " Datos");
//            
//            if (listadatosxTipoProducto.size() > 0) {
//                indxgrado2++;
//                p = new Paragraph();
//                p.setFont(fh2);
//                p.setAlignment(justificado);
//                p.add(indxgrado1 + "." + indxgrado2 + ". Estudio de solicitudes por " + jerarquiaProducto.get(i).get("PRODUCTO").replaceAll("_", " ") + ":\n\n");
//                documento.add(p);
//
//                List<Map<String, String>> listadocentexTipoProducto = data_list(1, listadatosxTipoProducto, new String[]{"No. IDENTIFICACION"});
//                indxgrado3 = 0;
//                for (int j = 0; j < listadocentexTipoProducto.size(); j++) {
//                    List<Map<String, String>> listadatosdocentexTipoProducto = data_list(3, listadatosxTipoProducto, new String[]{"No. IDENTIFICACION<->" + listadocentexTipoProducto.get(j).get("No. IDENTIFICACION")});
//                    indxgrado3 += 1;
//                    p = new Paragraph();
//                    p.setFont(fh3);
//                    p.setAlignment(justificado);
//                    c = new Chunk(indxgrado1 + "." + indxgrado2 + "." + indxgrado3 + ". " + listadocentexTipoProducto.get(j).get("NOMBRE DOCENTE") + ":\n", fh2);
//                    p.add(c);
//                    c = new Chunk("Identificación: ", fh2);
//                    p.add(c);
//                    p.add(listadocentexTipoProducto.get(j).get("No. IDENTIFICACION") + "\n" + listadocentexTipoProducto.get(j).get("FACULTAD") + ".\n");
//                    c = new Chunk("Categoría: ", fh2);
//                    p.add(c);
//                    p.add(listadocentexTipoProducto.get(j).get("CATEGORIA_DOCENTE") + " - " + listadocentexTipoProducto.get(j).get("TIPO VINCULACION") + " desde el " + fechaEnletras(listadocentexTipoProducto.get(j).get("FECHA INGRESO"), 0) + ".\n");
//                    c = new Chunk("Tipo de solicitud: ", fh2);
//                    p.add(c);
//                    p.add("Puntos por la publicación de " + listadatosdocentexTipoProducto.size() + " " + jerarquiaProducto.get(i).get("PRODUCTO").replaceAll("_", " "
//                            + "") + ". (" + jerarquiaProducto.get(i).get("NORMA") + ").\n");
//                    //documento.add(p);
//
//                    for (int k = 0; k < listadatosdocentexTipoProducto.size(); k++) {
//                        //p= new Paragraph();
//                        p.setFont(fh3);
//                        p.setAlignment(justificado);
//                        c = new Chunk("Soportes " + (listadatosdocentexTipoProducto.size() > 1 ? ((k + 1) + " " + jerarquiaProducto.get(i).get("PRODUCTO").replaceAll("_", " ") + ": ") : ": "), fh2);
//                        p.add(c);
//                        String soportes = getSoportes(listadatosdocentexTipoProducto, k);
//                        p.add(soportes + ".\n");
//                        //documento.add(p);
//                    }
//
//                    String decision = getDecision(listadatosdocentexTipoProducto);
//                    //p= new Paragraph();
//                    p.setFont(fh3);
//                    p.setAlignment(justificado);
//                    c = new Chunk("Decisión: ", fh2);
//                    p.add(c);
//                    p.add("Revisada la documentación y después de analizar lo establecido en las normas, el Comité decide " + decision + "\n");
//                    //documento.add(p);
//
//                    if (!listadocentexTipoProducto.get(j).get("NOTA").equals("N/A")) {
//                        p.setFont(fh3);
//                        p.setAlignment(justificado);
//                        c = new Chunk("Nota: ", fh2);
//                        p.add(c);
//                        p.add(listadocentexTipoProducto.get(j).get("NOTA") + "\n");
//
//                    }
//
//                    documento.add(p);
//
//                }
//            }
//        }
//        
//        if (listaProposiciones.size() > 0) {
//            indxgrado1++;
//            indxgrado2 = 0;
//            p = new Paragraph();
//            p.setFont(fh3);
//            p.setAlignment(justificado);
//            c = new Chunk(indxgrado1 + ". PROPOSICIONES Y VARIOS: \n\n", fh2);
//            p.add(c);
//            for (int y = 0; y < listaProposiciones.size(); y++) {
//                indxgrado2++;
//                p.add(indxgrado1 + "." + indxgrado2 + " " + listaProposiciones.get(y).get("NOMBRE SOLICITUD") + "\n");
//            }
//            documento.add(p);
//        }
//        indxgrado1++;
//        p = new Paragraph();
//        p.setFont(fh3);
//        p.setAlignment(justificado);
//        c = new Chunk(indxgrado1 + ". CIERRE: \n\n", fh2);
//        p.add(c);
//        p.add(" Siendo las " + listaDatosacta.get("HORA_FIN") + " se da por terminada la sesión");
//        documento.add(p);
////
//    }
//</editor-fold>
    
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
            System.out.println("ERROR data_list-->" + e.toString());
        }
        return rlista;
    }

    private String getDecision(List<Map<String, String>> listadatosdocentexTipoProducto)throws Exception{
        String respuesta = "";
        String respuestaxEstado = "";
        int banderasuma=0;
        double sumatoria_puntos = 0;
        String articulo ="";
        for (int f = 0; f < listadatosdocentexTipoProducto.size(); f++) {
            try {
                sumatoria_puntos += Double.parseDouble(ValidarNumero(listadatosdocentexTipoProducto.get(f).get("PUNTOS").replace(",", ".")));
            } catch (Exception ex) {
                Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                throw new Exception(" "+ex.getMessage());
            }
            if(listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Aprobado")){
                banderasuma=1;
            }
            System.out.println("SUMATORIA PUNTOS "+ sumatoria_puntos);
            articulo = getArticuloxTipoProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"));
            respuestaxEstado += (listadatosdocentexTipoProducto.size() > 1 ? " por "+articulo+" " 
                    + (getNombreNumero((f + 1), getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "ARTICULO")) + " " 
                                   +getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")+ " "
                                //+ listadatosdocentexTipoProducto.get(f).get("TIPO PRODUCTO").replaceAll("_", " ")
                                ) : "");
            if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Aprobado")) {
                if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                    
                    respuestaxEstado += "aprobar la promoción en el escalafón docente a la categoría " 
                                        + listadatosdocentexTipoProducto.get(f).get("NOMBRE_SOLICITUD") 
                                        + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")?
                                        " al":" a la")+" "
                                        + "docente de planta "
                                        + listadatosdocentexTipoProducto.get(f).get("NOMBRE_DOCENTE")
                                        + " y asignar " + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + " ("
                                        + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") " + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE");    
                    
                    if (listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD").equals(listaDatosacta.get("FECHA_ACTA"))||Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals("N/A")) {
                        respuestaxEstado += " a partir de la fecha de la presente sesión";
                    } else if(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD").length()>10){
                        respuestaxEstado +=  listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD") + ".";
                    }else{
                        respuestaxEstado += " a partir de " + fechaEnletras(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD"), 0) + "";
                    }
                    
                } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Ingreso_a_la_Carrera_Docente")) {
                    respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO").equals("Titulacion")) {
                    respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                } else {
                    //System.out.println("hello aqui voy?....." + listadatosdocentexTipoProducto.get(f).get("TIPO PUNTAJE"));
                    boolean cond = listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").equals("puntos salariales");
                    if (cond) {
//                        System.out.println("ENTRE?....." + listadatosdocentexTipoProducto.get(f).get("TIPO PUNTAJE"));
                        respuestaxEstado += "asignarle "
                                    + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")?
                                    "al":"a la")+" "
                                    + "docente "
                                    + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS"))
                                    + " (" + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") " 
                                    + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE") ;
                        
                        if (listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD").equals(listaDatosacta.get("FECHA_ACTA"))||Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD")).equals("N/A"))  {
                            respuestaxEstado += " a partir de la fecha de la presente sesión";
                        }else if(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD").length()>10){
                            respuestaxEstado += " "+ listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD") + ".";
                        }else {
                            respuestaxEstado += " a partir del " + fechaEnletras(listadatosdocentexTipoProducto.get(f).get("RETROACTIVIDAD"), 0) + "";
                        }

                        if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/A")) {
                            respuestaxEstado += " considerando que "+articulo+" " 
                                    +getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                                    //+ listadatosdocentexTipoProducto.get(f).get("TIPO PRODUCTO").replaceAll("_", " ")
                                    
                                    + " corresponde a un(a) " 
                                    + (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")).equals("N/A")?
                                        "producto":
                                        Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO"))+" ")
                                    //+ listadatosdocentexTipoProducto.get(f).get("SUBTIPO PRODUCTO")
                                    + (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                    ? " de carácter " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL"))
                                    : "")
                                    + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")) + ") ";
                        }
                        try {
                            if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) > 3) {
                                if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) < 6) {
                                    respuestaxEstado += " y teniendo en cuenta el número de autores (literal b; numeral III, artículo 10 del Decreto 1279 de 2002).";
                                } else {
                                    respuestaxEstado += " y teniendo en cuenta el número de autores (literal c; numeral III, artículo 10 del Decreto 1279 de 2002).";
                                }
                            }
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(" "+ex.getMessage());
                        }

                    } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                        respuestaxEstado += "reconocer "
                                + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")?
                                    "al":"a la")+" "
                                +"docente "
                                + getNumeroDecimal(listadatosdocentexTipoProducto.get(f).get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("PUNTOS")) + ") " + listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE") + ".";

                        if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/D")  && !Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")).equals("#N/A")) {
                            respuestaxEstado += " considerando que "+articulo+" " 
                                    +getDatoJerarquiaProducto(listadatosdocentexTipoProducto.get(f).get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                                    //+ listadatosdocentexTipoProducto.get(f).get("TIPO PRODUCTO").replaceAll("_", " ")
                                    + " corresponde a un(a) " 
                                    + (Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO")).equals("N/A")?
                                        "producto":
                                        Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("SUBTIPO_PRODUCTO"))+" ")
                                    //+ listadatosdocentexTipoProducto.get(f).get("SUBTIPO PRODUCTO")
                                    + (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                    ? " de carácter " + listadatosdocentexTipoProducto.get(f).get("NACIONAL/INTERNACIONAL/REGIONAL")
                                    : "")
                                    + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("NORMA")) + ") ";
                        }
                        try {
                            if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) > 3) {
                                if (Integer.parseInt(ValidarNumeroDec(listadatosdocentexTipoProducto.get(f).get("No._AUTORES"))) < 6) {
                                    respuestaxEstado += " y teniendo en cuenta el número de autores (literal b; numeral I, artículo 21 del Decreto 1279 de 2002).";
                                } else {
                                    respuestaxEstado += " y teniendo en cuenta el número de autores (literal c; numeral I, artículo 21 del Decreto 1279 de 2002).";
                                }
                            }
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionActas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(" "+ex.getMessage());
                        }

                    } else if (listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").trim().equals("no aplica")||listadatosdocentexTipoProducto.get(f).get("TIPO_PUNTAJE").trim().equals("convalidacion")) {
                        respuestaxEstado += " "+Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
                    }
                }
            } else if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Rechazado")) {
                respuestaxEstado += "no dar trámite a la solicitud "
                        + (listadatosdocentexTipoProducto.get(f).get("SEXO").toUpperCase().equals("M")?
                            "del":"de la")+" "
                        +"docente en razón a que " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
            } else if (listadatosdocentexTipoProducto.get(f).get("RESPUESTA_CIARP").equals("Enviar a Pares")) {
                respuestaxEstado += "enviar el producto a revisión por parte de pares externos de Colciencias teniendo en cuenta lo establecido en el Artículo 15 del Decreto 1279 del 2002";
            } else {
                respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(f).get("DECISION"));
            }

        }

        String add = (listadatosdocentexTipoProducto.size() > 1
                ? " Para un total de "+getNumeroDecimal(""+formateador.format(sumatoria_puntos))+" (" + formateador.format(sumatoria_puntos) + ") " + 
                listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE") 
                + " por la productividad presentada"
                : "");
        System.out.println(" SUMATORIA QUE SALE **** " +listadatosdocentexTipoProducto.get(0).get("NOMBRE_DOCENTE")+" " +sumatoria_puntos);
        

        respuesta = respuestaxEstado;
        if (banderasuma==1) {
            respuesta += ((!listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon _Docente")
                    && !listadatosdocentexTipoProducto.get(0).get("TIPO_PRODUCTO").equals("Titulacion")
                    && listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE").equals("puntos salariales"))
                    ? add + ", de conformidad con lo establecido en el Numeral 22 del Artículo Primero del Acuerdo de Seguimiento N° 001 de 2004 y al Parágrafo III del Artículo 12 del Decreto 1279 de 2002"
                    : listadatosdocentexTipoProducto.get(0).get("TIPO_PUNTAJE").equals("puntos de bonificacion")? add : "");
            System.out.println(" ADDD a ver que tiene "+ add+" "+ listadatosdocentexTipoProducto.get(0).get("NOMBRE_DOCENTE") );
        }
        
//        else if (listadatosdocentexTipoProducto.get(0).get("RESPUESTA CIARP").equals("Aprobado")&& listadatosdocentexTipoProducto.get(0).get("TIPO PUNTAJE").equals("puntos de bonificacion")){
//            respuesta+= add;
//    }
        return respuesta;
    }

    private String getSoportes(List<Map<String, String>> listadatosdocentexTipoProducto, int k) {
        String respuestaSoporte = "";
        String datosProducto = "";
        String datosSoporte = "";
        String repl ="";
        System.out.println("------ENTRE------------- " + listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO"));
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
                        + "); " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia del artículo "+ datosProducto+"; "+datosSoporte;
                break;
            case "Libro":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia de la libro "+ datosProducto+"; "+datosSoporte;
                break;
            case "Capitulo_de_Libro":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia del capitulo "+ datosProducto+"; "+datosSoporte;
                break;
            case "Ponencias_en_Eventos_Especializados":
                datosSoporte =ComillasSoporte( Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; en el " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia de la ponencia "+ datosProducto+"; "+datosSoporte;
                break;
            case "Publicaciones_Impresas_Universitarias":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia de la publicación impresa universitaria "+ datosProducto+"; "+datosSoporte;
                break;
            case "Reseñas_Críticas":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de la " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("PUBLINDEX"))
                        + "); " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia de la reseña "+ datosProducto+"; "+datosSoporte;
                break;
            case "Traducciones":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                        + "\"; de " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("ISSN/ISBN"))
                        + "; " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("PUBLINDEX"))
                        + "); " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                repl = getReplaceSoporteArticulo(datosSoporte);
                respuestaSoporte = "Copia de la traducción "+ datosProducto+"; "+datosSoporte;
                break;
            case "Direccion_de_Tesis":
                datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD")) + "\" ";
//                
                respuestaSoporte = "Copia del acta de sustentación de la "+ datosProducto +"; "+ datosSoporte;
                break;
            default:
                if (!Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("No._AUTORES")).equals( "N/A")) {
                    datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                            + "\" " + listadatosdocentexTipoProducto.get(k).get("No._AUTORES") + " autor(es)";
//                    repl = getReplaceSoporteArticulo(datosSoporte);
                    respuestaSoporte = "Copia de "+listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO") + datosProducto +", "+ datosSoporte;
                } else {
                    datosSoporte = ComillasSoporte(Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("SOPORTES")));
                    datosProducto = " " + Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("NOMBRE_SOLICITUD"))
                            + " ";
//                    repl = getReplaceSoporteArticulo(datosSoporte);
                    respuestaSoporte = "Copia de " +Utilidades.Utilidades.decodificarElemento(listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO")) + datosProducto+ ", "+ datosSoporte;
    //                respuestaSoporte = datosSoporte.replaceFirst("Copia de", "Copia de " + listadatosdocentexTipoProducto.get(k).get("TIPO PRODUCTO").replaceAll("_", " ") + " " + datosProducto);
                }
                break;
        }

        return respuestaSoporte;
    }

    private String getNumeroDecimal(String numero) {
        String retorno = "";
        System.out.println("numero----->"+numero);
        if(numero.indexOf(",") > -1){
            
            numero = numero.replace(",", ".");
        }
        
        if (numero.indexOf(".") > -1) {
            System.out.println("numero------>"+numero);
            Double dat = Double.parseDouble(numero);
            System.out.println("dat------>"+dat);
            DecimalFormat df = new DecimalFormat("#.0");
            System.out.println("df.format(dat)------>"+df.format(dat));
            numero = df.format(dat);
            System.out.println("numero------>"+numero);
            numero = numero.replace(".", ",");
            System.out.println("numero final------>"+numero);
        }
        
        if (!numero.equals("N/A")) {
            if (numero.indexOf(",") > -1) {
                String[] numrs = numero.replace(",", "::").split("::");
                retorno = numeroEnLetras(Integer.parseInt(numrs[0].equals("")?"0":numrs[0]));
                if(Integer.parseInt(numrs[1]) > 0){
                    retorno += " coma ";
                    retorno += numeroEnLetras(Integer.parseInt(numrs[1]));// + " (" + numero + ")";
                }
            } else {
                retorno = numeroEnLetras(Integer.parseInt(numero));// + "(" + numero + ")";
            }
        }

        return retorno;
    }

    private String numeroEnLetras(int numero) {
        String[] Unidades, Decenas, Centenas;
        String Resultado = "";

        /**
         * ************************************************
         * Nombre de los números
         * ************************************************
         */
        Unidades = new String[]{"", "Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veintiún", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve"};
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
            Resultado = "Un Millón" + agregado;
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
         * Nombre de los números
         * ************************************************
         */
        Unidades = new String[]{"", "Primero", "Segundo", "Tercero", "Cuarto", "Quinto", "Sexto", "Séptimo", "Octavo", "Noveno", "Décimo", "Undécimo", "Duodécimo"};
        Decenas = new String[]{"","Decimo", "Vigésimo", "Trigésimo", "Cuadragésimo", "Quincuagésimo", "Sexagésimo", "Septuagésimo", "Octogésimo", "Nonagésimo"};
        Centenas = new String[]{"", "Centésimo", "Ducentésimo", "Tricentésimo", "Cuadringentésimo", "Quingentésimo", "Sexcentésimo", "Septingentésimo", "Octingentésimo", "Noningentésimo"};

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
//        } else if (numero >= 1000 && numero <= 1999) {
//            String agregado = "";
//            if (numero % 1000 != 0) {
//                agregado = " " + numeroOrdinales(numero % 1000);
//            } else {
//                agregado = "";
//            }
//            Resultado = "Mil" + agregado;
//        } else if (numero >= 2000 && numero <= 999999) {
//            String agregado = "";
//            if (numero % 1000 != 0) {
//                agregado = " " + numeroOrdinales(numero % 1000);
//            } else {
//                agregado = "";
//            }
//            Resultado = numeroEnLetras(numero / 1000) + " Mil" + agregado;
//        } else if (numero >= 1000000 && numero <= 1999999) {
//            String agregado = "";
//            if (numero % 1000000 != 0) {
//                agregado = " " + numeroEnLetras(numero % 1000000);
//            } else {
//                agregado = "";
//            }
//            Resultado = "Un Millón" + agregado;
//        } else if (numero >= 2000000 && numero <= 1999999999) {
//            String agregado = "";
//            if (numero % 1000000 != 0) {
//                agregado = " " + numeroEnLetras(numero % 1000000);
//            } else {
//                agregado = "";
//            }
//            Resultado = numeroEnLetras(numero / 1000000) + " Millones" + agregado;
        }
        return Resultado.toLowerCase();
    }
    
    private String ValidarNumero(String numero) throws Exception{
        return (numero.equals("N/A") ? "0" : numero);
    }

    private String fechaEnletras(String fecha, int opc) throws Exception {// 7/08/2012

        String fechaletra = "";
        if (!fecha.equals("N/A")) {
            try{
            String[] dividirFecha = fecha.split("/");
            String[] meses = {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"};
            System.out.println(" FECHA QUE LLEGA  "+fecha);
            String dia = numeroEnLetras(Integer.parseInt(dividirFecha[0]));
            String mes = meses[Integer.parseInt(dividirFecha[1]) - 1];

            fechaletra = dividirFecha[0] + " de " + mes + " de " + dividirFecha[2];
            if (opc == 1) {
                fechaletra = dia + " (" + dividirFecha[0] + ") días del mes de " + mes + " de " + dividirFecha[2];
            }
            }catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception(" "+ex.getMessage());
            }
        
        }
        return fechaletra;
    }    
    private String getSoportesColciencias(List<Map<String, String>> listadocentexArticuloColciencias, int k) {
        
        String soportes = Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("NOMBRE_SOLICITUD"))
                +"; de la "+Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("REVISTA/EVENTO/EDITORIAL"))+"; ISSN:"
                        +listadocentexArticuloColciencias.get(k).get("ISSN/ISBN")+"; "+Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("FECHA_PUBLICACION/REALIZACION"))+"; ("+Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(k).get("PUBLINDEX"))
                        +"); "+listadocentexArticuloColciencias.get(k).get("No._AUTORES")+"autor(es). \n";
    return soportes;
           
    }

    private String getDecisionColciencias(List<Map<String, String>> listadocentexArticuloColciencias) throws Exception{
       String respuesta = "";
        String respuestaxEstado = "";
        double sumatoria_puntos = 0;
        for (int f = 0; f < listadocentexArticuloColciencias.size(); f++) {
            try{
            sumatoria_puntos += Double.parseDouble(ValidarNumero(listadocentexArticuloColciencias.get(f).get("PUNTOS").replace(",", ".")));
            }
            catch(Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception(" "+ex.getMessage());
            }

            respuestaxEstado += (listadocentexArticuloColciencias.size() > 1 ? " por el " + (getNombreNumero((f + 1),"el")+ " artículo" ) : "");
            if (listadocentexArticuloColciencias.get(f).get("RESPUESTA_CIARP").equals("Aprobado")) {
                      
                   
                            respuestaxEstado += " asignarle "
                                    + (listadocentexArticuloColciencias.get(f).get("SEXO").toUpperCase().equals("M")?
                                    "al":"a la")+" "
                                    +"docente "
                                    + getNumeroDecimal(listadocentexArticuloColciencias.get(f).get("PUNTOS"))
                                    + " (" + listadocentexArticuloColciencias.get(f).get("PUNTOS") + ") " + listadocentexArticuloColciencias.get(f).get("TIPO_PUNTAJE");
                            if(listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD").length() > 10){
                                respuestaxEstado += listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD") + ".";
                            }else{
                                respuestaxEstado += " a partir de " + fechaEnletras(listadocentexArticuloColciencias.get(f).get("RETROACTIVIDAD"), 0) + ".";
                            }
                        

                        if (!Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("NORMA")).equals("#N/D")  && !Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("NORMA")).equals("#N/A")) {
                            respuestaxEstado += " considerando que el artículo" 
                                    + " corresponde a un(a) " + listadocentexArticuloColciencias.get(f).get("SUBTIPO_PRODUCTO")
                                    + "(" + listadocentexArticuloColciencias.get(f).get("NORMA") + ")";
                        }
                        if (Integer.parseInt(ValidarNumero(listadocentexArticuloColciencias.get(f).get("No._AUTORES"))) > 3) {
                            if (Integer.parseInt(ValidarNumero(listadocentexArticuloColciencias.get(f).get("No._AUTORES"))) < 6) {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal b; numeral III, artículo 10 del Decreto 1279 de 2002).";
                            } else {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal c; numeral III, artículo 10 del Decreto 1279 de 2002).";
                            }
                        }

                        } else if (listadocentexArticuloColciencias.get(f).get("RESPUESTA_CIARP").equals("Rechazado")) {
                respuestaxEstado +=  Utilidades.Utilidades.decodificarElemento(listadocentexArticuloColciencias.get(f).get("DECISION"));
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
        
        for(Map<String, String> obj: jerarquiaProducto){
            if(obj.get("PRODUCTO").equals(tipo)){
                ret = obj.get("ARTICULO");
                break;
            }
        }
        
        return ret;
    }
       
    private String getDatoJerarquiaProducto(String datoB, String KeyB, String KeyR) {
        String ret = "";
        
        for(Map<String, String> obj: jerarquiaProducto){
            if(obj.get(KeyB).equals(datoB)){
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
        
        if("LA".equals(articulo.toUpperCase())){
            nombre = nombre.substring(0, nombre.length()-1)+"a";
        }else{
            if(numero == 1 || numero == 3){
                nombre = nombre.substring(0, nombre.length()-1);
            }
        }
        
        return nombre;
    }
    
    private String FormatoCedula(String cedula){
        if(cedula.indexOf(",")>=0){
        String[] Cedul = cedula.split(",");
        cedula=Cedul[0];
        }
        String ret = "";
        int suma = 0;
        for(int i = cedula.length()-1; i >= 0; i--){
            if(suma == 3){
                suma = 0;
                ret = "."+ret;
            }
            suma++;
            ret = ""+cedula.charAt(i)+ret;
        }
        
        return ret;
    }

//    private String getReplaceSoporteArticulo(String datosSoporte) {
//        String ret = "";
//        
//        if(listadatosdocentexTipoProducto.get(k).get("TIPO_PRODUCTO").equ){
//            ret = "articulo";
//        }else if(datosSoporte.indexOf("artículo")>-1){
//            ret = "artículo";
//        }else if(datosSoporte.indexOf("Articulo")>-1){
//            ret = "Articulo";
//        }else if(datosSoporte.indexOf("Artículo")>-1){
//            ret = "Artículo";
//        }else if(datosSoporte.indexOf("ARTICULO")>-1){
//            ret = "ARTICULO";
//        }else if(datosSoporte.indexOf("ARTÍCULO")>-1){
//            ret = "ARTÍCULO";
//        }else if(datosSoporte.indexOf("Libro")>-1){
//            ret = "Libro";
//        }else if(datosSoporte.indexOf("libro")>-1){
//            ret = "libro";
//        }else if(datosSoporte.indexOf("LIBRO")>-1){
//            ret = "LIBRO";
//        }else if(datosSoporte.indexOf("capitulo de libro")>-1){
//            ret = "capitulo de libro";
//        }else if(datosSoporte.indexOf("capítulo de libro")>-1){
//            ret = "capítulo de libro";
//        }else if(datosSoporte.indexOf("Capítulo de libro")>-1){
//            ret = "Capítulo de libro";
//        }else if(datosSoporte.indexOf("Capitulo de libro")>-1){
//            ret = "Capitulo de libro";
//        }else if(datosSoporte.indexOf("Capitulo De Libro")>-1){
//            ret = "Capitulo De Libro";
//        }else if(datosSoporte.indexOf("Capítulo De Libro")>-1){
//            ret = "Capítulo De Libro";
//        }else if(datosSoporte.indexOf("Capitulo de Libro")>-1){
//            ret = "Capitulo de Libro";
//        }else if(datosSoporte.indexOf("Capítulo de Libro")>-1){
//            ret = "Capítulo de Libro";
//        }else if(datosSoporte.indexOf("CAPITULO DE LIBRO")>-1){
//            ret = "CAPITULO DE LIBRO";
//        }else if(datosSoporte.indexOf("CAPÍTULO DE LIBRO")>-1){
//            ret = "CAPÍTULO DE LIBRO";
//        }else if(datosSoporte.indexOf("PONENCIA")>-1){
//            ret = "PONENCIA";
//        }else if(datosSoporte.indexOf("ponencia")>-1){
//            ret = "ponencia";
//        }else if(datosSoporte.indexOf("Ponencia")>-1){
//            ret = "Ponencia";
//        }else if(datosSoporte.indexOf("impresa universitaria")>-1){
//            ret = "impresa universitaria";
//        }else if(datosSoporte.indexOf("Impresa universitaria")>-1){
//            ret = "Impresa universitaria";
//        }else if(datosSoporte.indexOf("impresa Universitaria")>-1){
//            ret = "impresa Universitaria";
//        }else if(datosSoporte.indexOf("Impresa Universitaria")>-1){
//            ret = "Impresa Universitaria";
//        }else if(datosSoporte.indexOf("IMPRESA UNIVERSITARIA")>-1){
//            ret = "IMPRESA UNIVERSITARIA";
//        }else if(datosSoporte.indexOf("RESEÑA CRITICA")>-1){
//            ret = "RESEÑA CRITICA";
//        }else if(datosSoporte.indexOf("RESEÑA CRÍTICA")>-1){
//            ret = "RESEÑA CRÍTICA";
//        }else if(datosSoporte.indexOf("reseña critica")>-1){
//            ret = "reseña critica";
//        }else if(datosSoporte.indexOf("reseña crítica")>-1){
//            ret = "reseña crítica";
//        }else if(datosSoporte.indexOf("Reseña Crítica")>-1){
//            ret = "Reseña Crítica";
//        }else if(datosSoporte.indexOf("Reseña crítica")>-1){
//            ret = "Reseña crítica";
//        }else if(datosSoporte.indexOf("Reseña critica")>-1){
//            ret = "Reseña critica";
//        }else if(datosSoporte.indexOf("reseña Crítica")>-1){
//            ret = "reseña Crítica";
//        }else if(datosSoporte.indexOf("reseña Critica")>-1){
//            ret = "reseña Critica";
//        }else if(datosSoporte.indexOf("TRADUCCION DEL ARTICULO")>-1){
//            ret = "TRADUCCION DEL ARTICULO";
//        }else if(datosSoporte.indexOf("TRADUCCIÓN DEL ARTÍCULO")>-1){
//            ret = "TRADUCCIÓN DEL ARTÍCULO";
//        }else if(datosSoporte.indexOf("TRADUCCIÓN DEL ARTICULO")>-1){
//            ret = "TRADUCCIÓN DEL ARTICULO";
//        }else if(datosSoporte.indexOf("TRADUCCION DEL ARTÍCULO")>-1){
//            ret = "TRADUCCION DEL ARTÍCULO";
//        }else if(datosSoporte.indexOf("traduccion del articulo")>-1){
//            ret = "traduccion del articulo";
//        }else if(datosSoporte.indexOf("traducción del articulo")>-1){
//            ret = "traducción del articulo";
//        }else if(datosSoporte.indexOf("traduccion del artículo")>-1){
//            ret = "traduccion del artículo";
//        }else if(datosSoporte.indexOf("traducción del artículo")>-1){
//            ret = "traducción del artículo";
//        }else if(datosSoporte.indexOf("Traducción del Artículo")>-1){
//            ret = "Traducción del Artículo";
//        }else if(datosSoporte.indexOf("Traduccion del Articulo")>-1){
//            ret = "Traduccion del Articulo";
//        }else if(datosSoporte.indexOf("Traducción del artículo")>-1){
//            ret = "Traducción del artículo";
//        }else if(datosSoporte.indexOf("Traduccion del articulo")>-1){
//            ret = "Traduccion del articulo";
//        }else if(datosSoporte.indexOf("de sustentación")>-1){
//            ret = "de sustentación";
//        }else if(datosSoporte.indexOf("de Sustentación")>-1){
//            ret = "de Sustentación";
//        }else if(datosSoporte.indexOf("de sustentacion")>-1){
//            ret = "de sustentacion";
//        }else if(datosSoporte.indexOf("de Sustentacion")>-1){
//            ret = "de Sustentacion";
//        }else if(datosSoporte.indexOf("Copia de")>-1){
//            ret = "Copia de";
//        }else if(datosSoporte.indexOf("copia de")>-1){
//            ret = "copia de";
//        }
//        
//        
//        
//        
//        
//        return ret;
//    }
   
    public String ValidarNumeroDec(String valor){
//        String ret = formateador.format(""+valor);
//        System.out.println("************************ValidarNumeroDec****************************");
//        System.out.println("valorform---->"+ret);
//        if(valor.indexOf(".")>-1){
//            String[] dat = valor.replace(".", ":").split(":");
//            if(dat[1].equals("00")){
//                ret = dat[0];
//            }else{
//                ret = ret.replace(".", ",");
//            }
//        }
//        System.out.println("ret--------->"+ret);
        String retorno = "";
        System.out.println("numero ini----->"+valor);
        if(valor.indexOf(",") > -1){
            
            valor = valor.replace(",", ".");
        }
        else{
            retorno =valor;
        }
        
        if (valor.indexOf(".") > -1) {
            
            System.out.println("numero------>"+valor);
            Double dat = Double.parseDouble(valor);
            System.out.println("dat------>"+dat);
            DecimalFormat df = new DecimalFormat("0.0");
            System.out.println("df.format(dat)------>"+df.format(dat));
            valor = df.format(dat);
            System.out.println("numero formateado------>"+valor);
            valor = valor.replace(".", ",");
            String[] daot= valor.split(",");
            System.out.println("DAOT [0]" +daot[0]);
//            if (daot[0].equals("")){
//                daot[0]= "0";
//                valor = daot[0]+valor;
//                System.out.println("DAO[0] = "+daot[0]);
//                System.out.println("DAO[1] = "+daot[1]);
//            }
            if(daot[1].equals("0")){
                retorno = daot[0];
            }else{
                retorno = valor;
            }
            System.out.println("numero final------>"+retorno);
        }
        
//        if (!numero.equals("N/A")) {
//            if (numero.indexOf(",") > -1) {
//                String[] numrs = numero.replace(",", "::").split("::");
//                retorno = numeroEnLetras(Integer.parseInt(numrs[0].equals("")?"0":numrs[0]));
//                if(Integer.parseInt(numrs[1]) > 0){
//                    retorno += " coma ";
//                    retorno += numeroEnLetras(Integer.parseInt(numrs[1]));// + " (" + numero + ")";
//                }
//            } else {
//                retorno = numeroEnLetras(Integer.parseInt(numero));// + "(" + numero + ")";
//            }
//        }

        return retorno;
 
//        return ret;
    }
    
}
