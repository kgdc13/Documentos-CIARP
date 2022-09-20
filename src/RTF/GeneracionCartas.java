/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package RTF;

//import Configuracion.SystemParam;
import Excel.ControlArchivoExcel;
import Excel.gestorInformes;
import static RTF.GeneracionActas.listaDatosacta;
import static RTF.GeneracionResoluciones.URL;
import static RTF.GeneracionResoluciones.lineaDeTexto;
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
import com.lowagie.text.rtf.headerfooter.RtfHeaderFooterGroup;
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
import java.util.Calendar;

/**
 *
 * @author rjulio
 */
public class GeneracionCartas {
    
//    static String URL = "C:\\CIARP\\Cartas_Sesión_.rtf";
    static String tipoProducto = "Ingreso a la Carrera Docente";
    static String cedula = "";
    static int bandera = 0;
    static int banderaTP = 0;
    DecimalFormat formateador = new DecimalFormat("#.#");
    static List<Map<String, String>> jerarquiaProducto = new ArrayList<>();
    static Map<String, String> listaDatosacta = new HashMap<>();
    static Map<String, String> Datosacta = new HashMap<>();
    ArrayList<Map<String, String>> datos2 = new ArrayList<>();
    static Map<String, String> DatosCartas = new HashMap<>();
    static ArrayList<ArrayList<Map<String, String>>> DatosNumeralesCartas = new ArrayList<ArrayList<Map<String, String>>>();
    public List<Map<String, String>> listaDatos = new ArrayList<>();
    public String ruta;
    public String consecutivo;
    public String anio;
    public int indxgrado1 = 0;
    public int indxgrado2 = 0;
    public int indxgrado3 = 0;
    Calendar jc;
    int dia;
    int mes;
    int fanio;
    String fecha;
    
    public GeneracionCartas() {
        URL = "C:\\CIARP\\Cartas_Sesion_" + listaDatosacta.get("No ACTA") + ".rtf";
        
        tipoProducto = "Ingreso a la Carrera Docente";
        cedula = "";
        bandera = 0;
        banderaTP = 0;
        jerarquiaProducto = new ArrayList<>();
        listaDatosacta = new HashMap<>();
        Datosacta = new HashMap<>();
        DatosNumeralesCartas = new ArrayList<ArrayList<Map<String, String>>>();
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
    
    public void LeerArchivo() {

        this.ruta = ruta;
        this.consecutivo = consecutivo;
        ControlArchivo contArchivo = new ControlArchivo(ruta);
        contArchivo.LeerArchivo();
        BufferedReader br = contArchivo.getBuferDeLectura();
        String lineaDeTexto;
        

    }
    
    public Map<String, String> Generarcartas(String ruta, String consecutivo, String anio) throws DocumentException, IOException, Exception {
        ControlArchivoExcel con = new ControlArchivoExcel();
        BaseFont calibriFont = BaseFont.createFont("C:\\windows\\Fonts\\CALIBRI.TTF", "Cp1252", true);
        BaseFont arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
        System.out.println("¨*****************Generarcartas****************" + ruta);
        Map<String, String> respuesta = new HashMap<>();
        this.ruta = ruta;
        this.consecutivo = consecutivo;
        this.anio = anio;
        int numeral = 1;
        ControlArchivo contArchivo = new ControlArchivo(ruta);
        contArchivo.LeerArchivo();
        BufferedReader br = contArchivo.getBuferDeLectura();
        String lineaDeTexto;
//        LeerArchivo();
  //<editor-fold defaultstate="collapsed" desc="Lectura Orden del Día">
        String extP = ruta.substring(ruta.lastIndexOf(".") + 1);
        System.out.println("************************EMPIEZA LECTURA DE ORDEN DEL DÍA");
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

        //<editor-fold defaultstate="collapsed" desc="FUENTES Y ESTILO">
        Font fh1 = new Font(arialFont);
        fh1.setColor(Color.BLACK);
        fh1.setSize(11);
        fh1.setStyle("bold");
//                fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("underlined");
        
        Font fh3 = new Font(arialFont);
        fh3.setSize(11);
        fh3.setColor(Color.BLACK);
        
        Font fh3b = new Font(arialFont);
        fh3b.setSize(11);
        fh3b.setColor(Color.BLACK);
        fh3b.setStyle("bold");
        
        Font fh3c = new Font(arialFont);
        fh3c.setSize(8);
        fh3c.setColor(Color.BLACK);
        fh3c.setStyle("italic");
        
        Font fh4 = new Font(arialFont);
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");
        //fh2.setStyle("bold");

        Font af10 = new Font(arialFont);
        af10.setSize(10);
        af10.setColor(Color.BLACK);
        
        Font af10b = new Font(arialFont);
        af10b.setSize(10);
        af10b.setColor(Color.BLACK);
        af10b.setStyle("bold");
        
        Font af7 = new Font(arialFont);
        af7.setSize(7);
        af7.setColor(Color.BLACK);
        
        Font af7b = new Font(arialFont);
        af7b.setSize(7);
        af7b.setColor(Color.BLACK);
        af7b.setStyle("bold");
        
        Paragraph p = new Paragraph();
        int justificado = Paragraph.ALIGN_JUSTIFIED;
        int centrado = Paragraph.ALIGN_CENTER;
        int left = Paragraph.ALIGN_LEFT;
        int right = Paragraph.ALIGN_RIGHT;
        //</editor-fold>
        System.out.println("Comienza Generacion");
        List<Map<String, String>> listaDocentes = data_list(1, listaDatos, new String[]{"No._IDENTIFICACION"});
        String encode = Encode();
        Map<String, String> datos1 = new HashMap<>();
        String asunto = "";
        Document documento = new Document();
        documento = new Document(PageSize.LETTER);
        documento.setMargins(70, 70, 69, 55);
        URL = "C:\\CIARP\\Cartas_sesión_" + listaDatosacta.get("No_ACTA") + ".rtf";
        RtfWriter2.getInstance(documento, new FileOutputStream(URL));
        documento.open();
        System.out.println("*********URLSSSS**************URL " + URL);
        int conseAdd = Integer.parseInt(consecutivo);
        int banderar = 0, bandPtsSal = 0, bandPtsBon = 0;
        System.out.println("listaDocentes.size ===" + listaDocentes.size());
        for (int j = 0; j < listaDocentes.size(); j++) {
            System.out.println("*********************RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR****************************************Docente N°" + j + " - " + listaDocentes.get(j).get("NOMBRE_DOCENTE"));
            banderar = 0;
            bandPtsBon = 0;
            bandPtsSal = 0;
            DatosCartas.put("NOMBRE_ARCHIVO", "correspondencia_");
            DatosCartas.put("NUMACTA", "" + listaDatosacta.get("No_ACTA"));
            jc = Calendar.getInstance();
            dia = jc.get(Calendar.DATE);
            mes = jc.get(Calendar.MONTH) + 1;
            fanio = jc.get(Calendar.YEAR);
            fecha = "" + dia + "/" + "" + mes + "/" + "" + fanio;
            
            
            if (!listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Proposiciones_y_varios")) {

//            //<editor-fold defaultstate="collapsed" desc="HEADER AND FOOTER">
//            Table headerTable;
//            Table headerTableTxt;
//
//            System.out.println("$$$$$$$$$$$$$$$$ PAGINA N°" + j);
//            Image imgE = Image.getInstance("C:\\CIARP\\encabezado.png");
//
//            headerTable = new Table(1, 1);
//            headerTable.setAlignment(Cell.ALIGN_RIGHT);
//            headerTable.setWidth(90);
//
//            Cell celda = new Cell(imgE);
//            celda.setBorder(0);
//            celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
//            headerTable.addCell(celda);
//
//
//            Image imgF = Image.getInstance("C:\\CIARP\\footer.png");
//            Table footertable = new Table(1, 1);
//            footertable.setWidth(90);
//            footertable.setAlignment(Cell.ALIGN_RIGHT);
//
//
//            celda = new Cell(imgF);
//            celda.setBorder(0);
//            celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
//
//            footertable.addCell(celda);
//            RtfHeaderFooter header = new RtfHeaderFooter(headerTable);
//            RtfHeaderFooter footer = new RtfHeaderFooter(footertable);
//
//            documento.setHeader(header);
//            documento.setFooter(footer);
//
//            //</editor-fold>
                Table headerTable;
                Image imgE = Image.getInstance("C:\\CIARP\\encabezado.png");
                
                headerTable = new Table(1, 1);
                headerTable.setAlignment(Cell.ALIGN_RIGHT);
                headerTable.setWidth(90);
                
                Cell celda = new Cell(imgE);
                celda.setBorder(0);
                celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
                headerTable.addCell(celda);
                documento.add(headerTable);
                
                p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(fh3);
                p.add("Santa Marta, " + fechaEnletras(fecha, 0));
                documento.add(p);
                
                if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh1);
                    p.add("CIARP-" + "N/A" + "-" + anio + "\n\n");
                    documento.add(p);
                } else {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh1);
                    p.add("CIARP-" + conseAdd + "-" + anio + "\n\n");
                    documento.add(p);
                }
                
                p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(fh3);
                p.add("Docente \n");
                Chunk c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("NOMBRE_DOCENTE")), fh3b);
                p.add(c);
                p.add("\n"+Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("CORREO")));
                System.out.println("########################################"+listaDocentes.get(j).get("CORREO"));
                p.add("\n" + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("FACULTAD")) + "\n");
                p.add("Universidad del Magdalena \n");
                documento.add(p);
                
                String docente = Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("NOMBRE_DOCENTE"));
                if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Titulacion")) {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh3);
                    c = new Chunk("ASUNTO: ", fh3b);
                    p.add(c);
                    p.add("Respuesta a solicitud de puntos por titulación. \n");
                    documento.add(p);
                    asunto = "Respuesta a solicitud de puntos por titulación";
                } else if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh3);
                    c = new Chunk("ASUNTO: ", fh3b);
                    p.add(c);
                    p.add("Respuesta a solicitud de ascenso en el escalafón docente. \n");
                    documento.add(p);
                    asunto = "Respuesta a solicitud de ascenso en el escalafón docente";
                } else if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Revision_de_la_correspondencia")) {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh3);
                    c = new Chunk("ASUNTO: ", fh3b);
                    p.add(c);
                    p.add("Respuesta a comunicación enviada \n");
                    documento.add(p);
                    asunto = "Respuesta a comunicación enviada";
                } else {
                    p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(fh3);
                    c = new Chunk("ASUNTO: ", fh3b);
                    p.add(c);
                    p.add("Respuesta a solicitud de puntos por productividad académica. \n");
                    documento.add(p);
                    asunto = "Respuesta a solicitud de puntos por productividad académica";
                }
                
                p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(fh3);
                p.add("Cordial saludo, \n");
                documento.add(p);
                
                List<Map<String, String>> listaProductosxDocentes = data_list(3, listaDatos, new String[]{"No._IDENTIFICACION<->" + listaDocentes.get(j).get("No._IDENTIFICACION")});
                int numsolicitudes = listaProductosxDocentes.size();
                System.out.println("&&&&&&&&& lista productos docentes " + listaProductosxDocentes.size());
                System.out.println("listaDocentes.get(j).get(\"TIPO_PRODUCTO\")--->" + listaDocentes.get(j).get("TIPO_PRODUCTO"));
                System.out.println("ACTAS DATOSSS" + listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Titulacion")) {
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(fh3);
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                    try{
                    p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                    }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                    
                    p.add(" estudió su solicitud de puntos por el título de " + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("NOMBRE_SOLICITUD")) + " \n");
                    p.add(" Una vez revisada la documentación y verificado el cumplimiento de la norma el Comité determinó " + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("DECISION")) + " \n");
                    documento.add(p);
                    
                    if (listaDocentes.get(j).get("RESPUESTA_CIARP").equals("Aprobado")) {
                        banderar = 1;
                        System.out.println("BANDERA DERNTO IF " + banderar);
                    }
                    System.out.println(" BANDERAAAAAA " + banderar);
                    if (banderar == 1) {
                        p = new Paragraph(10);
                        p.setAlignment(justificado);
                        p.setFont(fh3);
                        p.add("No obstante, se indica que el Rector de la Universidad "
                                + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                                + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                                + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
                        documento.add(p);
                    }
//<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                    datos1 = new HashMap<>();
                    System.out.print("CONSECUTIVO::------------------" + conseAdd);
                    datos1.put("CONSECUTIVO", " " + conseAdd);
                    datos1.put("ASUNTO", asunto);
                    datos1.put("DIRIGIDO_A", docente);
                    datos1.put("FECHA", fecha);
                    datos2.add(datos1);
                    //</editor-fold>
                    conseAdd++;
                } else if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                    String retroactividad = "";
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(fh3);
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                   try{
                    p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                   }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                    
                    p.add(" estudió la solicitud de ascenso " + (listaDocentes.get(j).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " docente de planta " + listaDocentes.get(j).get("NOMBRE_DOCENTE") + ", de la categoría " + listaDocentes.get(j).get("CATEGORIA_DOCENTE")
                            + " a " + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("NOMBRE_SOLICITUD")) + " \n \n");
                    try {
                        p.add("Después de revisar el cumplimiento de los requisitos, el Comité decidió aprobar la promoción en el escalafón " + (listaDocentes.get(j).get("SEXO").equals("M") ? "del " : "de la ") + "docente y asignarle " + getNumeroDecimal(listaDocentes.get(j).get("PUNTOS")) + " (" + ValidarNumeroDec(listaDocentes.get(j).get("PUNTOS")) + ") puntos salariales");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        respuesta.put("ESTADO", "ERROR");
                        respuesta.put("MENSAJE", ""+ex.getMessage());
                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                        }else{
                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                        }
                        return respuesta;
                    }
                    if (Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")).equals("N/A")) {
                        System.out.print("ESTO EN IF DE RETROATIVIDAD " + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")) + "///////" + listaDatosacta.get("FECHA_ACTA"));
                        retroactividad += " a partir de la fecha de la presente sesión.";
                    } else if (Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")).length() > 10) {
                        retroactividad += Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")) + ".";
                    } else {
                        try{
                        retroactividad += " a partir de " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("RETROACTIVIDAD")), 0) + ".";
                        }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                    }
                    p.add(retroactividad + "\n");
                    documento.add(p);
                    
                    if (listaDocentes.get(j).get("RESPUESTA_CIARP").equals("Aprobado")) {
                        banderar = 1;
                        System.out.println("BANDERA DERNTO IF " + banderar);
                    }
                    System.out.println(" BANDERAAAAAA " + banderar);
                    if (banderar == 1) {
                        p = new Paragraph(10);
                        p.setAlignment(justificado);
                        p.setFont(fh3);
                        p.add("No obstante, se indica que el Rector de la Universidad "
                                + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                                + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                                + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
                        documento.add(p);
                    }
                } else if (listaDocentes.get(j).get("TIPO_PRODUCTO").equals("Revision_de_la_correspondencia")) {
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(fh3);
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                    try{
                    p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                    }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                    p.add(" estudió su comunicación en la cual manifiesta:" + " \n");
                    documento.add(p);
                    
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setIndentationLeft(13);
                    p.setIndentationRight(12);
                    p.setFont(fh3c);
                    try{
                    String[] carta = Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("NOMBRE_SOLICITUD")).split(":");
                        System.out.println(" CARTA "+carta[0]);
                        
                    p.add(carta[1] + "\n");
                    }catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        respuesta.put("ESTADO", "ERROR");
                        respuesta.put("MENSAJE", ""+ex.getMessage());
                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                        }else{
                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                        }
                        return respuesta;
                    }
                    documento.add(p);
                    
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(fh3);
                    p.add("" + Utilidades.Utilidades.decodificarElemento(listaDocentes.get(j).get("DECISION")) + "\n");
                    documento.add(p);

                    //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                    datos1 = new HashMap<>();
                    datos1.put("CONSECUTIVO", " " + conseAdd);
                    datos1.put("ASUNTO", asunto);
                    datos1.put("DIRIGIDO_A", docente);
                    datos1.put("FECHA", fecha);
                    datos2.add(datos1);
                    //</editor-fold>
                    conseAdd++;
                } else {
                    double puntos_salariales = 0;
                    double puntos_bonificacion = 0;
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(fh3);
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                    try{
                    p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                    }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                    
                    p.add(" estudió " + (listaProductosxDocentes.size() > 1 ? "sus solicitudes de puntos por los productos: " : "su solicitud de puntos por el producto:") + " \n");
                    documento.add(p);
                    
                    for (int i = 0; i < jerarquiaProducto.size(); i++) {
                        List<Map<String, String>> listaProductosxJerarquia = data_list(3, listaProductosxDocentes, new String[]{"TIPO_PRODUCTO<->" + jerarquiaProducto.get(i).get("PRODUCTO")});
                        System.out.println("&&&&&&&&&&&&&&&&&&&&&& productosss " + jerarquiaProducto.get(i).get("PRODUCTO"));
                        System.out.println("***************size LPXJER " + listaProductosxJerarquia.size() + " valor i" + i);
                        
                        if (listaProductosxJerarquia.size() > 0) {
                            System.out.println("lista productos jerarquia-->" + listaProductosxJerarquia.size());
                            for (int l = 0; l < listaProductosxJerarquia.size(); l++) {
                                System.out.println("***************size LPXJER  en for " + listaProductosxJerarquia.size() + "NOMBRE PRODU" + listaProductosxJerarquia.get(l).get("NOMBRE_SOLICITUD"));
                                System.out.println("*******l-->" + l + "******+");
                                System.out.println("tipo-puntaje--" + listaProductosxJerarquia.get(l).get("TIPO_PUNTAJE"));
                                p = new Paragraph(10);
                                p.setAlignment(justificado);
                                p.setFont(fh3);
                                String nameproduct = getNOMBREPRODUCTO(listaProductosxJerarquia.get(l));
                                System.out.println("name--->" + nameproduct);
                                System.out.println(" PRODUCTO ___ BUSCANDO EL ARTICULO" + listaDocentes.get(l).get("TIPO_PRODUCTO"));                                
                                p.add("• " + listaProductosxJerarquia.get(l).get("TIPO_PRODUCTO").replace("_", " ") + ":" + nameproduct + "\n");
                                try {
                                    p.add("Por este producto el Comité determinó " + Utilidades.Utilidades.decodificarElemento(getDecision(listaProductosxJerarquia.get(l))) + "\n ");
                                } catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                        if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
                                documento.add(p);
                                System.out.println(" AQUI LA DECISION ES :::::: " + listaProductosxJerarquia.get(l).get("RESPUESTA_CIARP") + " Y LA BANDERA " + banderar);
                                if (listaProductosxJerarquia.get(l).get("TIPO_PUNTAJE").equals("puntos salariales")) {
                                    try {
                                        puntos_salariales += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get(l).get("PUNTOS").replace(",", ".")));
                                    } catch (Exception ex) {
                                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                         if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;
                                    }
                                    bandPtsSal = 1;
                                }
                                
                                if (listaProductosxJerarquia.get(l).get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                                    try {
                                        puntos_bonificacion += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get(l).get("PUNTOS").replace(",", ".")));
                                    } catch (Exception ex) {
                                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                         if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                            respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                          }else{
                                            respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                            }
                                        return respuesta;
                                    }
                                    bandPtsBon =1;
                                }
                                if (listaProductosxJerarquia.get(l).get("RESPUESTA_CIARP").equals("Aprobado")) {
                                    banderar = 1;
                                    System.out.println("BANDERA DERNTO IF " + banderar);
                                }
                                
                            }
                        }
                        
                    }
                    System.out.println(" BANDERAAAAAA " + banderar);
                    if (banderar == 1) {
                        System.out.println(" PUNTOS SALARIALES " + puntos_salariales);
                        System.out.println(" PUNTOS DE BONIFICACION  " + puntos_bonificacion);
                        
                        if (bandPtsSal == 1) {
                            p = new Paragraph(10);
                            p.setAlignment(justificado);
                            p.setFont(fh3);
                            try {
                                p.add("Para un total de " + getNumeroDecimal(Double.toString(puntos_salariales)) + " (" + ValidarNumeroDec("" + puntos_salariales) + ") puntos salariales por la productividad presentada. \n ");
                            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", "Posible error de digitación"+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                    }
                                return respuesta;
                            }
                            documento.add(p);
                        }
                        System.out.println("condicion- bonificacion-"+(listaDocentes.get(j).get("TIPO_PUNTAJE").equals("puntos de bonificacion")));
                        if (bandPtsBon == 1) {
                            
                            p = new Paragraph(10);
                            p.setAlignment(justificado);
                            p.setFont(fh3);
                            
                            try {
                                p.add("Para un total de " + getNumeroDecimal(Double.toString(puntos_bonificacion)) + " (" + ValidarNumeroDec("" + puntos_bonificacion) + ") puntos de bonificación por la productividad presentada. \n ");
                            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+docente);
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD"));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+listaDocentes.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                    }
                                return respuesta;
                            }
                            documento.add(p);
                        }
                        p = new Paragraph(10);
                        p.setAlignment(justificado);
                        p.setFont(fh3);
                        p.add("No obstante, se indica que el Rector de la Universidad "
                                + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                                + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                                + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
                        documento.add(p);
                    }
                    //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
                    datos1 = new HashMap<>();
                    datos1.put("CONSECUTIVO", " " + conseAdd);
                    datos1.put("ASUNTO", asunto);
                    datos1.put("DIRIGIDO_A", docente);
                    datos1.put("FECHA", fecha);
                    datos2.add(datos1);
                    //</editor-fold>
                    conseAdd++;

                }


                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(fh3);
                p.add("Agradezco su atención sobre el particular. \n\n");
                p.add("Atentamente, \n \n");
                c = new Chunk("OSCAR HUMBERTO GARCIA VARGAS \n", fh3b);
                p.add(c);
                p.add("Vicerrector Académico \n");
                p.add("Presidente Comité Interno de Asignación y Reconocimiento de Puntaje \n");
                documento.add(p);
                
                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(af7);
                p.add("Con Copia: Hoja de Vida" + (listaDocentes.get(j).get("SEXO").equals("M") ? " del " : " de la ") + "Docente – Dirección de Talento Humano ");
                documento.add(p);
                Image imgF = Image.getInstance("C:\\CIARP\\footer.png");
                Table footertable = new Table(1, 1);
                footertable.setWidth(90);
                footertable.setAlignment(Cell.ALIGN_RIGHT);
                
                celda = new Cell(imgF);
                celda.setBorder(0);
                celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
                
                footertable.addCell(celda);
                documento.add(footertable);
                
                documento.newPage();
            }
        }
        System.out.println("-----------------DOCUMENTO CERRADO----------------");
        documento.close();
        System.out.println("***********************URL " + URL);
        File f = new File(URL);
        f.createNewFile();
        respuesta.put("ESTADO", "OK");
        respuesta.put("MENSAJE", "Las cartas se generaron satisfactoriamente.");
        int result = JOptionPane.showConfirmDialog(null, "¿Desea abrir el documento?");
        if (result == JOptionPane.YES_OPTION) {
            Desktop.getDesktop().open(f);
        }
        DatosNumeralesCartas.add(datos2);
        System.out.println("DatosResoluciones-->" + DatosCartas.size());
        System.out.println("DatosNumeralesResoluciones-->" + DatosNumeralesCartas.size());
        gestorInformes gi = new gestorInformes(DatosCartas, DatosNumeralesCartas);
        gi.iniciar();
        
        System.out.println("documento.getPageNumber()--->" + documento);
//          
        return respuesta;
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
            System.out.println("ERROR data_list-->" + e.toString());
        }
        return rlista;
    }
    
    private String getDecision(Map<String, String> listaProductosxJerarquia) throws Exception{
        String respuesta = "";
        String respuestaxEstado = "";
        double sumatoria_puntos = 0;
        int banderasuma = 0;
        String articulo = "";
        try {
         
            sumatoria_puntos += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get("PUNTOS").replace(",", ".")));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception(" "+ex.getMessage());
        }
        banderasuma = 1;
        articulo = getArticuloxTipoProducto(listaProductosxJerarquia.get("TIPO_PRODUCTO"));
//        
        if (listaProductosxJerarquia.get("RESPUESTA_CIARP").equals("Aprobado")) {
            if (listaProductosxJerarquia.get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                
                try {
                    respuestaxEstado += "aprobar la promoción en el escalafón docente a la categoría "
                            + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NOMBRE_SOLICITUD"))
                            + (listaProductosxJerarquia.get("SEXO").toUpperCase().equals("M")
                            ? " al" : " a la") + " "
                            + "docente de planta "
                            + listaProductosxJerarquia.get("NOMBRE_DOCENTE")
                            + " y asignar " + getNumeroDecimal(listaProductosxJerarquia.get("PUNTOS")) + " ("
                            + ValidarNumeroDec(listaProductosxJerarquia.get("PUNTOS")) + ") " + listaProductosxJerarquia.get("TIPO_PUNTAJE");
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                    throw new Exception(" "+ex.getMessage());
                    
                }
                
                if (listaProductosxJerarquia.get("RETROACTIVIDAD").equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals("N/A")) {
                    respuestaxEstado += " a partir de la fecha de la presente sesión.";
                } else if (listaProductosxJerarquia.get("RETROACTIVIDAD").length() > 10) {
                    respuestaxEstado += listaProductosxJerarquia.get("RETROACTIVIDAD") + ".";
                } else {
                    try{
                    respuestaxEstado += " a partir de " + fechaEnletras(listaProductosxJerarquia.get("RETROACTIVIDAD"), 0) + ".";
                        } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                    }
                }
                    
                    
                
            } else if (listaProductosxJerarquia.get("TIPO_PRODUCTO").equals("Ingreso_a_la_Carrera_Docente")) {
                respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
            } else if (listaProductosxJerarquia.get("TIPO_PRODUCTO").equals("Titulacion")) {
                respuestaxEstado += Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
            } else {
              
                boolean cond = listaProductosxJerarquia.get("TIPO_PUNTAJE").equals("puntos salariales");
                if (cond) {
                    try {
                     
                        respuestaxEstado += "asignarle "
                                + getNumeroDecimal(listaProductosxJerarquia.get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listaProductosxJerarquia.get("PUNTOS")) + ") "
                                + listaProductosxJerarquia.get("TIPO_PUNTAJE");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                    }
                    
                    if (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals("N/A")) {
                        respuestaxEstado += " a partir de la fecha de la presente sesión";
                    } else if (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).length() > 10) {
                        respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")) + ".";
                    } else {
                        try {
                        respuestaxEstado += " a partir del " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")), 0) + ".";
                        }
                        catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                        }
                        
                    }
                    
                    if (!Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NORMA")).equals("#N/A")) {
                        respuestaxEstado += " considerando que " + articulo + " "
                                + getDatoJerarquiaProducto(listaProductosxJerarquia.get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                           

                                + " corresponde a un(a) "
                                + (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("SUBTIPO_PRODUCTO")).equals("N/A")
                                ? "producto "
                                : Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("SUBTIPO_PRODUCTO")))
                            
                                + (!Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                ? " de carácter " + listaProductosxJerarquia.get("NACIONAL/INTERNACIONAL/REGIONAL")
                                : "")
                                + " (" + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NORMA")) + ") ";
                    }
                    System.out.println("NO AUTORES "+listaProductosxJerarquia.get("No._AUTORES"));
                    try {
                        if (Integer.parseInt(ValidarNumero(ValidarNumeroDec(listaProductosxJerarquia.get("No._AUTORES")))) > 3) {
                            if (Integer.parseInt(ValidarNumero(ValidarNumeroDec(listaProductosxJerarquia.get("No._AUTORES")))) < 6) {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal b; numeral III, artículo 10 del Decreto 1279 de 2002).";
                            } else {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal c; numeral III, artículo 10 del Decreto 1279 de 2002).";
                            }
                        }
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                    }
                    
                } else if (listaProductosxJerarquia.get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                    try {
                        respuestaxEstado += "reconocer "
                                + getNumeroDecimal(listaProductosxJerarquia.get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listaProductosxJerarquia.get("PUNTOS")) + ") " + listaProductosxJerarquia.get("TIPO_PUNTAJE") + ".";
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                    }
                    
                    if (!Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NORMA")).equals("#N/D") && !listaProductosxJerarquia.get("NORMA").equals("#N/A")) {
                        respuestaxEstado += " considerando que " + articulo + " "
                                + getDatoJerarquiaProducto(listaProductosxJerarquia.get("TIPO_PRODUCTO"), "PRODUCTO", "NPRODUCTO")
                              
                                + " corresponde a un(a) "
                                + (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("SUBTIPO_PRODUCTO")).equals("N/A")
                                ? "producto "
                                : Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("SUBTIPO_PRODUCTO")))
                              
                                + (!Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NACIONAL/INTERNACIONAL/REGIONAL")).equals("N/A")
                                ? " de carácter " + listaProductosxJerarquia.get("NACIONAL/INTERNACIONAL/REGIONAL")
                                : "")
                                + " (" + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("NORMA")) + ") ";
                    }
                    System.out.println("listaProductosxJerarquia.get(f).get(\"No_AUTORES\")--" + listaProductosxJerarquia.get("No._AUTORES"));
                    try {
                        
                        if (Integer.parseInt(ValidarNumero(ValidarNumeroDec(listaProductosxJerarquia.get("No._AUTORES")))) > 3) {
                            
                            if (Integer.parseInt(ValidarNumero(ValidarNumeroDec(listaProductosxJerarquia.get("No._AUTORES")))) < 6) {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal b; numeral I, artículo 21 del Decreto 1279 de 2002).";
                            } else {
                                respuestaxEstado += " y teniendo en cuenta el número de autores (literal c; numeral I, artículo 21 del Decreto 1279 de 2002).";
                            }
                        }
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" "+ex.getMessage());
                    }
                    
                } else if (listaProductosxJerarquia.get("TIPO_PUNTAJE").trim().equals("no aplica") || listaProductosxJerarquia.get("TIPO_PUNTAJE").trim().equals("convalidacion")) {
                    respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
                }
            }
        } else if (listaProductosxJerarquia.get("RESPUESTA_CIARP").equals("Rechazado")) {
            respuestaxEstado += "no dar trámite a su solicitud "
                    + "en razón a que " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
        } else if (listaProductosxJerarquia.get("RESPUESTA_CIARP").equals("Enviar a Pares")) {
            respuestaxEstado += "enviar el producto a revisión por parte de pares externos de Colciencias teniendo en cuenta lo establecido en el Artículo 15 del Decreto 1279 del 2002";
        } else {
            respuestaxEstado += listaProductosxJerarquia.get("DECISION");
        }


        respuesta = respuestaxEstado;
        return respuesta;
    }
    
    
    
    private String getNumeroDecimal(String numero) {
        String retorno = "";
        System.out.println("numero----->" + numero);
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
        Decenas = new String[]{"", "Decimo", "Vigésimo", "Trigésimo", "Cuadragésimo", "Quincuagésimo", "Sexagésimo", "Septuagésimo", "Octogésimo", "Nonagésimo"};
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

        }
        return Resultado.toLowerCase();
    }
    
    private String ValidarNumero(String numero) throws Exception{
        return (numero.equals("N/A") ? "0" : numero);
    }
    
    private String fechaEnletras(String fecha, int opc) throws Exception{// 7/08/2012

        String fechaletra = "";
        if (!fecha.equals("N/A")) {
            String[] dividirFecha = fecha.split("/");
            System.out.println("FECHA DIVIDIDA " + fecha);
            String[] meses = {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"};
            System.out.println("TAMAÑO FECHA DIVIFDIR" + dividirFecha.length);
            String dia = numeroEnLetras(Integer.parseInt(dividirFecha[0]));
            String mes = meses[Integer.parseInt(dividirFecha[1]) - 1];
            
            fechaletra = dividirFecha[0] + " de " + mes + " de " + dividirFecha[2];
            if (opc == 1) {
                fechaletra = dia + " (" + dividirFecha[0] + ") días del mes de " + mes + " de " + dividirFecha[2];
            }
        }
        
        return fechaletra;
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
    
    
    private String getReplaceSoporteArticulo(String datosSoporte) {
        String ret = "";
        
        if (datosSoporte.indexOf("articulo") > -1) {
            ret = "articulo";
        } else if (datosSoporte.indexOf("artículo") > -1) {
            ret = "artículo";
        } else if (datosSoporte.indexOf("Articulo") > -1) {
            ret = "Articulo";
        } else if (datosSoporte.indexOf("Artículo") > -1) {
            ret = "Artículo";
        } else if (datosSoporte.indexOf("ARTICULO") > -1) {
            ret = "ARTICULO";
        } else if (datosSoporte.indexOf("ARTÍCULO") > -1) {
            ret = "ARTÍCULO";
        } else if (datosSoporte.indexOf("Libro") > -1) {
            ret = "Libro";
        } else if (datosSoporte.indexOf("libro") > -1) {
            ret = "libro";
        } else if (datosSoporte.indexOf("LIBRO") > -1) {
            ret = "LIBRO";
        } else if (datosSoporte.indexOf("capitulo de libro") > -1) {
            ret = "capitulo de libro";
        } else if (datosSoporte.indexOf("capítulo de libro") > -1) {
            ret = "capítulo de libro";
        } else if (datosSoporte.indexOf("Capítulo de libro") > -1) {
            ret = "Capítulo de libro";
        } else if (datosSoporte.indexOf("Capitulo de libro") > -1) {
            ret = "Capitulo de libro";
        } else if (datosSoporte.indexOf("Capitulo De Libro") > -1) {
            ret = "Capitulo De Libro";
        } else if (datosSoporte.indexOf("Capítulo De Libro") > -1) {
            ret = "Capítulo De Libro";
        } else if (datosSoporte.indexOf("Capitulo de Libro") > -1) {
            ret = "Capitulo de Libro";
        } else if (datosSoporte.indexOf("Capítulo de Libro") > -1) {
            ret = "Capítulo de Libro";
        } else if (datosSoporte.indexOf("CAPITULO DE LIBRO") > -1) {
            ret = "CAPITULO DE LIBRO";
        } else if (datosSoporte.indexOf("CAPÍTULO DE LIBRO") > -1) {
            ret = "CAPÍTULO DE LIBRO";
        } else if (datosSoporte.indexOf("PONENCIA") > -1) {
            ret = "PONENCIA";
        } else if (datosSoporte.indexOf("ponencia") > -1) {
            ret = "ponencia";
        } else if (datosSoporte.indexOf("Ponencia") > -1) {
            ret = "Ponencia";
        } else if (datosSoporte.indexOf("impresa universitaria") > -1) {
            ret = "impresa universitaria";
        } else if (datosSoporte.indexOf("Impresa universitaria") > -1) {
            ret = "Impresa universitaria";
        } else if (datosSoporte.indexOf("impresa Universitaria") > -1) {
            ret = "impresa Universitaria";
        } else if (datosSoporte.indexOf("Impresa Universitaria") > -1) {
            ret = "Impresa Universitaria";
        } else if (datosSoporte.indexOf("IMPRESA UNIVERSITARIA") > -1) {
            ret = "IMPRESA UNIVERSITARIA";
        } else if (datosSoporte.indexOf("RESEÑA CRITICA") > -1) {
            ret = "RESEÑA CRITICA";
        } else if (datosSoporte.indexOf("RESEÑA CRÍTICA") > -1) {
            ret = "RESEÑA CRÍTICA";
        } else if (datosSoporte.indexOf("reseña critica") > -1) {
            ret = "reseña critica";
        } else if (datosSoporte.indexOf("reseña crítica") > -1) {
            ret = "reseña crítica";
        } else if (datosSoporte.indexOf("Reseña Crítica") > -1) {
            ret = "Reseña Crítica";
        } else if (datosSoporte.indexOf("Reseña crítica") > -1) {
            ret = "Reseña crítica";
        } else if (datosSoporte.indexOf("Reseña critica") > -1) {
            ret = "Reseña critica";
        } else if (datosSoporte.indexOf("reseña Crítica") > -1) {
            ret = "reseña Crítica";
        } else if (datosSoporte.indexOf("reseña Critica") > -1) {
            ret = "reseña Critica";
        } else if (datosSoporte.indexOf("TRADUCCION DEL ARTICULO") > -1) {
            ret = "TRADUCCION DEL ARTICULO";
        } else if (datosSoporte.indexOf("TRADUCCIÓN DEL ARTÍCULO") > -1) {
            ret = "TRADUCCIÓN DEL ARTÍCULO";
        } else if (datosSoporte.indexOf("TRADUCCIÓN DEL ARTICULO") > -1) {
            ret = "TRADUCCIÓN DEL ARTICULO";
        } else if (datosSoporte.indexOf("TRADUCCION DEL ARTÍCULO") > -1) {
            ret = "TRADUCCION DEL ARTÍCULO";
        } else if (datosSoporte.indexOf("traduccion del articulo") > -1) {
            ret = "traduccion del articulo";
        } else if (datosSoporte.indexOf("traducción del articulo") > -1) {
            ret = "traducción del articulo";
        } else if (datosSoporte.indexOf("traduccion del artículo") > -1) {
            ret = "traduccion del artículo";
        } else if (datosSoporte.indexOf("traducción del artículo") > -1) {
            ret = "traducción del artículo";
        } else if (datosSoporte.indexOf("Traducción del Artículo") > -1) {
            ret = "Traducción del Artículo";
        } else if (datosSoporte.indexOf("Traduccion del Articulo") > -1) {
            ret = "Traduccion del Articulo";
        } else if (datosSoporte.indexOf("Traducción del artículo") > -1) {
            ret = "Traducción del artículo";
        } else if (datosSoporte.indexOf("Traduccion del articulo") > -1) {
            ret = "Traduccion del articulo";
        } else if (datosSoporte.indexOf("de sustentación") > -1) {
            ret = "de sustentación";
        } else if (datosSoporte.indexOf("de Sustentación") > -1) {
            ret = "de Sustentación";
        } else if (datosSoporte.indexOf("de sustentacion") > -1) {
            ret = "de sustentacion";
        } else if (datosSoporte.indexOf("de Sustentacion") > -1) {
            ret = "de Sustentacion";
        } else if (datosSoporte.indexOf("Copia de") > -1) {
            ret = "Copia de";
        } else if (datosSoporte.indexOf("copia de") > -1) {
            ret = "copia de";
        }
        
        return ret;
    }
    
    public String ValidarNumeroDec(String valor) throws Exception{

        String retorno = "";
        System.out.println("numero----->" + valor);
        if (valor.indexOf(",") > -1) {
            
            valor = valor.replace(",", ".");
        } else {
            retorno = valor;
        }
        
        if (valor.indexOf(".") > -1) {
            
            System.out.println("numero------>" + valor);
            Double dat = Double.parseDouble(valor);
            System.out.println("dat------>" + dat);
            DecimalFormat df = new DecimalFormat("0.0");
            System.out.println("df.format(dat)------>" + df.format(dat));
            valor = df.format(dat);
            System.out.println("numero------>" + valor);
            valor = valor.replace(".", ",");
            String[] daot = valor.split(",");
            System.out.println("DAOT [0]" + daot[0]);

            if (daot[1].equals("0")) {
                retorno = daot[0];
            } else {
                retorno = valor;
            }
            System.out.println("numero final------>" + retorno);
        }


        return retorno;


    }
    
    private String getNOMBREPRODUCTO(Map<String, String> datos) {
        String datosProducto = "";
        try {
            switch (datos.get("TIPO_PRODUCTO")) {

                case "Articulo":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Art_Col":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Libro":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Capitulo_de_Libro":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Ponencias_en_Eventos_Especializados":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; en el " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + "; " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Publicaciones_Impresas_Universitarias":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la  " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + "; " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Reseñas_Críticas":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la  " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Traducciones":
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    break;
                case "Direccion_de_Tesis":
                    datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "") + " ";
                    break;
                default:
                    if (!Utilidades.Utilidades.decodificarElemento(datos.get("No._AUTORES")).equals("N/A")) {
                        datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                                + ", " + ValidarNumeroDec(datos.get("No._AUTORES")) + " autor(es).";
                    } else {
                        datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                                + " ";
                    }
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return datosProducto;
    }
}
