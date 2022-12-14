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
 * @author rjulio, rramos, kdelosreyes
 */
public class GeneracionCartas {

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
    BaseFont arialFont;
    Calendar jc;
    int dia;
    int mes;
    int fanio;
    String fecha;
    int conseAdd = 0;
    String asunto = "";
    Map<String, String> datos1 = new HashMap<>();

    Paragraph p = new Paragraph();
    int justificado = Paragraph.ALIGN_JUSTIFIED;
    int centrado = Paragraph.ALIGN_CENTER;
    int left = Paragraph.ALIGN_LEFT;
    int right = Paragraph.ALIGN_RIGHT;


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
        
        arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
        
        
        Map<String, String> respuesta = new HashMap<>();
        this.ruta = ruta;
        this.consecutivo = consecutivo;
        conseAdd = Integer.parseInt(this.consecutivo);
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
        System.out.println("rutaPuntos--->" + ruta);
        System.out.println("extP--->" + extP);
        if (extP.equals("xlsx")) {
            listaDatos = con.LeerExcelDesdeAct(ruta, 2, "ORDEN DEL DIA");
            listaDatosacta = con.LeerExcelParametrosAct(ruta, 1, 1, "ORDEN DEL DIA");
        } else {
            listaDatos = con.LeerExcelDesde(ruta, 2, "ORDEN DEL DIA");
            listaDatosacta = con.LeerExcelParametros(ruta, 1, 1);
        }

//</editor-fold>
        
        List<Map<String, String>> listaDocentes = data_list(1, listaDatos, new String[]{"No._IDENTIFICACION"});
        String encode = Encode();
        
        String asunto = "";
        Document documento = new Document();
        documento = new Document(PageSize.LETTER);
        documento.setMargins(70, 70, 69, 55);
        URL = "C:\\CIARP\\Cartas_sesión_" + listaDatosacta.get("No_ACTA") + ".rtf";
        RtfWriter2.getInstance(documento, new FileOutputStream(URL));
        documento.open();
        

        int banderar = 0, bandPtsSal = 0, bandPtsBon = 0;
        System.out.println("listaDocentes.size ===" + listaDocentes.size());
        for (int j = 0; j < listaDocentes.size(); j++) {
            List<Map<String, String>> listaProductosxDocentes = data_list(3, listaDatos, new String[]{"No._IDENTIFICACION<->" + listaDocentes.get(j).get("No._IDENTIFICACION")});
            
            banderar = 0;
            bandPtsBon = 0;
            bandPtsSal = 0;
            List<Map<String, String>> listaTipoCartas = getTipoCarta(listaProductosxDocentes);

            DatosCartas.put("NOMBRE_ARCHIVO", "correspondencia_");
            DatosCartas.put("NUMACTA", "" + listaDatosacta.get("No_ACTA"));
            jc = Calendar.getInstance();
            dia = jc.get(Calendar.DATE);
            mes = jc.get(Calendar.MONTH) + 1;
            fanio = jc.get(Calendar.YEAR);
            fecha = "" + dia + "/" + "" + mes + "/" + "" + fanio;

            for (Map<String, String> tipoCarta : listaTipoCartas) {
                List<Map<String, String>> datosProductosxCarta = getDatosCarta(tipoCarta.get("tipo"), listaProductosxDocentes);
                plantillaCartaEncabezado(documento, listaDatosacta, listaDocentes.get(j), tipoCarta.get("tipo"));
                
                if (!tipoCarta.get("tipo").equalsIgnoreCase("Proposiciones_y_varios")) {
                    
                    if (tipoCarta.get("tipo").equalsIgnoreCase("Ingreso_a_la_Carrera_Docente")) {
                        try {
                            GenerarCartaIngreso(documento, listaDocentes.get(j), datosProductosxCarta);
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            respuesta.put("ESTADO", "ERROR");
                            respuesta.put("MENSAJE", "" + ex.getMessage());
                            respuesta.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_DOCENTE")));
                            respuesta.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("TIPO_PRODUCTO")));
                            if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).length() <= 100) {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")));
                            } else {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).substring(0, 100));
                            }
                            return respuesta;
                        }
                    }
                    if (tipoCarta.get("tipo").equalsIgnoreCase("Revision_de_la_correspondencia")) {
                       try{
                        GenerarCartaRevision(documento, listaDocentes.get(j), datosProductosxCarta);
                       }catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            respuesta.put("ESTADO", "ERROR");
                            respuesta.put("MENSAJE", "" + ex.getMessage());
                            respuesta.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_DOCENTE")));
                            respuesta.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("TIPO_PRODUCTO")));
                            if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).length() <= 100) {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")));
                            } else {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).substring(0, 100));
                            }
                            return respuesta;
                        }
                    }
                    if (tipoCarta.get("tipo").equalsIgnoreCase("Ascenso_en_el_Escalafon_Docente")) {
                 try{
                        GenerarCartaAscenso(documento, listaDocentes.get(j), datosProductosxCarta);
                       }catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            respuesta.put("ESTADO", "ERROR");
                            respuesta.put("MENSAJE", "" + ex.getMessage());
                            respuesta.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_DOCENTE")));
                            respuesta.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("TIPO_PRODUCTO")));
                            if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).length() <= 100) {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")));
                            } else {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).substring(0, 100));
                            }
                            return respuesta;
                        }
                    }
                if (tipoCarta.get("tipo").equalsIgnoreCase("Titulacion")) {
                 try{
                        GenerarCartaTitulacion(documento, listaDocentes.get(j), datosProductosxCarta);
                       }catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            respuesta.put("ESTADO", "ERROR");
                            respuesta.put("MENSAJE", "" + ex.getMessage());
                            respuesta.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_DOCENTE")));
                            respuesta.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("TIPO_PRODUCTO")));
                            if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).length() <= 100) {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")));
                            } else {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).substring(0, 100));
                            }
                            return respuesta;
                        }
                    }
                    if (tipoCarta.get("tipo").equalsIgnoreCase("Productividad_Academica")) {
                        
                        try {
                            GenerarCartaProductividad(documento, listaDocentes.get(j), datosProductosxCarta);
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            respuesta.put("ESTADO", "ERROR");
                            respuesta.put("MENSAJE", "" + ex.getMessage());
                            respuesta.put("LINEA_ERROR_DOCENTE", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_DOCENTE")));
                            respuesta.put("LINEA_ERROR_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("TIPO_PRODUCTO")));
                            if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).length() <= 100) {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")));
                            } else {
                                respuesta.put("NOMBRE_PRODUCTO", "" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(j).get("NOMBRE_SOLICITUD")).substring(0, 100));
                            }
                            return respuesta;
                        }
                    }
                    plantillaCartaPiepagina(documento, listaDocentes.get(j));
                }
            
           }
       }
           
        
        documento.close();
        
        File f = new File(URL);
        f.createNewFile();
        respuesta.put("ESTADO", "OK");
        respuesta.put("MENSAJE", "Las cartas se generaron satisfactoriamente.");
        int result = JOptionPane.showConfirmDialog(null, "¿Desea abrir el documento?");
        if (result == JOptionPane.YES_OPTION) {
            Desktop.getDesktop().open(f);
        }
        DatosNumeralesCartas.add(datos2);
        
        gestorInformes gi = new gestorInformes(DatosCartas, DatosNumeralesCartas);
        gi.iniciar();

        
         
         
        return respuesta;
    }
    
    private void plantillaCartaEncabezado(Document documento, Map<String, String>listadatosacta, Map<String,String>listadocentes, String tipoCarta) throws DocumentException, IOException, Exception {

//            //<editor-fold defaultstate="collapsed" desc="Imagen de carta">
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
//            //</editor-fold>                
        Fonts f = new Fonts(arialFont);
        p = new Paragraph(10);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("Santa Marta, " + fechaEnletras(fecha, 0));
        documento.add(p);

        getConsecutivoCarta(documento, tipoCarta);

        
        String docente = Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        p = new Paragraph(10);
        p.setAlignment(left);
        
        Font d = f.SetFont(Color.black, 11,Fonts.NORMAL);
        p.setFont(d);
        p.add("Docente \n");
        Chunk c = new Chunk(docente, Fonts.SetFont(Color.black, 11,Fonts.BOLD));
        p.add(c);
        p.add("\n" + Utilidades.Utilidades.decodificarElemento(listadocentes.get("CORREO")));
        
        p.add("\n" + Utilidades.Utilidades.decodificarElemento(listadocentes.get("FACULTAD")) + "\n");
        p.add("Universidad del Magdalena \n");
        documento.add(p);

    }

   private void plantillaCartaPiepagina(Document documento, Map<String, String> listadocentes) throws DocumentException, IOException, Exception {
                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                p.add("Agradezco su atención sobre el particular. \n\n");
                p.add("Atentamente, \n \n");
               Chunk c = new Chunk("OSCAR HUMBERTO GARCIA VARGAS \n", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
                p.add(c);
                p.add("Vicerrector Académico \n");
                p.add("Presidente Comité Interno de Asignación y Reconocimiento de Puntaje \n");
                documento.add(p);

                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(Fonts.SetFont(Color.black, 7,Fonts.NORMAL));
                p.add("Con Copia: Hoja de Vida" + (listadocentes.get("SEXO").equals("M") ? " del " : " de la ") + "Docente – Dirección de Talento Humano ");
                documento.add(p);
                Image imgF = Image.getInstance("C:\\CIARP\\footer.png");
                Table footertable = new Table(1, 1);
                footertable.setWidth(90);
                footertable.setAlignment(Cell.ALIGN_RIGHT);

               Cell celda = new Cell(imgF);
                celda.setBorder(0);
                celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);

                footertable.addCell(celda);
                documento.add(footertable);

                documento.newPage();
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

    private String getDecision(Map<String, String> listaProductosxJerarquia) throws Exception {
        String respuesta = "";
        String respuestaxEstado = "";
        double sumatoria_puntos = 0;
        int banderasuma = 0;
        String articulo = "";
        try {

            sumatoria_puntos += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get("PUNTOS").replace(",", ".")));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception(" " + ex.getMessage());
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
                    throw new Exception(" " + ex.getMessage());

                }

                if (listaProductosxJerarquia.get("RETROACTIVIDAD").equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals("N/A")) {
                    respuestaxEstado += " a partir de la fecha de la presente sesión.";
                } else if (listaProductosxJerarquia.get("RETROACTIVIDAD").length() > 10) {
                    respuestaxEstado += listaProductosxJerarquia.get("RETROACTIVIDAD") + ".";
                } else {
                    try {
                        respuestaxEstado += " a partir de " + fechaEnletras(listaProductosxJerarquia.get("RETROACTIVIDAD"), 0) + ".";
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" " + ex.getMessage());
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
                        throw new Exception(" " + ex.getMessage());
                    }

                    if (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).equals("N/A")) {
                        respuestaxEstado += " a partir de la fecha de la presente sesión";
                    } else if (Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")).length() > 10) {
                        respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")) + ".";
                    } else {
                        try {
                            respuestaxEstado += " a partir del " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("RETROACTIVIDAD")), 0) + ".";
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(" " + ex.getMessage());
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
                        throw new Exception(" " + ex.getMessage());
                    }

                } else if (listaProductosxJerarquia.get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                    try {
                        respuestaxEstado += "reconocer "
                                + getNumeroDecimal(listaProductosxJerarquia.get("PUNTOS"))
                                + " (" + ValidarNumeroDec(listaProductosxJerarquia.get("PUNTOS")) + ") " + listaProductosxJerarquia.get("TIPO_PUNTAJE") + ".";
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception(" " + ex.getMessage());
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
                        throw new Exception(" " + ex.getMessage());
                    }

                } else if (listaProductosxJerarquia.get("TIPO_PUNTAJE").trim().equals("no aplica") || listaProductosxJerarquia.get("TIPO_PUNTAJE").trim().equals("convalidacion")) {
                    respuestaxEstado += " " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
                }
            }
        } else if (listaProductosxJerarquia.get("RESPUESTA_CIARP").equals("Rechazado")) {
            respuestaxEstado += "no dar trámite a su solicitud "
                    + "en razón a que " + Utilidades.Utilidades.decodificarElemento(listaProductosxJerarquia.get("DECISION"));
        } else if (listaProductosxJerarquia.get("RESPUESTA_CIARP").equals("Enviar a pares")) {
            respuestaxEstado += "enviar el producto a revisión por parte de pares externos de Colciencias teniendo en cuenta lo establecido en el Artículo 15 del Decreto 1279 del 2002";
        } else {
            respuestaxEstado += listaProductosxJerarquia.get("DECISION");
        }

        respuesta = respuestaxEstado;
        return respuesta;
    }

    private String getNumeroDecimal(String numero) {
        String retorno = "";
        
        if (numero.indexOf(",") > -1) {

            numero = numero.replace(",", ".");
        }

        if (numero.indexOf(".") > -1) {
           
            Double dat = Double.parseDouble(numero);
            
            DecimalFormat df = new DecimalFormat("#.0");
            
            numero = df.format(dat);
            
            numero = numero.replace(".", ",");
            
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

    private String ValidarNumero(String numero) throws Exception {
        return (numero.equals("N/A") ? "0" : numero);
    }

    private String fechaEnletras(String fecha, int opc) throws Exception {// 7/08/2012

        String fechaletra = "";
        if (!fecha.equals("N/A")) {
            String[] dividirFecha = fecha.split("/");
            
            String[] meses = {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"};
            
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

    public String ValidarNumeroDec(String valor) throws Exception {

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

    private List<Map<String, String>> getTipoCarta(List<Map<String, String>> datosxproducto) {
        List<Map<String, String>> listaTipoProducto = data_list(1, datosxproducto, new String[]{"TIPO_PRODUCTO"});
        List<Map<String, String>> listaTipoCarta = new ArrayList<Map<String, String>>();
        HashMap<String, String> nombreCarta;
        int banderaProductividad = 0;
        for (int i = 0; i < listaTipoProducto.size(); i++) {
            nombreCarta = new HashMap<>();
            if (listaTipoProducto.get(i).get("TIPO_PRODUCTO").equals("Titulacion")) {
                nombreCarta.put("tipo", "Titulacion");
            } else if (listaTipoProducto.get(i).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                nombreCarta.put("tipo", "Ascenso_en_el_Escalafon_Docente");
            } else if (listaTipoProducto.get(i).get("TIPO_PRODUCTO").equals("Revision_de_la_correspondencia")) {
                nombreCarta.put("tipo", "Revision_de_la_correspondencia");
            } else if (listaTipoProducto.get(i).get("TIPO_PRODUCTO").equals("Ingreso_a_la_Carrera_Docente")) {
                nombreCarta.put("tipo", "Ingreso_a_la_Carrera_Docente");
            } else if (listaTipoProducto.get(i).get("TIPO_PRODUCTO").equals("Proposiciones_y_varios")) {
                nombreCarta.put("tipo", "Proposiciones_y_varios");
            } else {
                if (banderaProductividad == 0) {
                    nombreCarta.put("tipo", "Productividad_Academica");
                    banderaProductividad = 1;
                }
            }
            if(nombreCarta.containsKey("tipo"))
                listaTipoCarta.add(nombreCarta);
        }
        return listaTipoCarta;
    }

    private List<Map<String, String>> getDatosCarta(String tipoCarta, List<Map<String, String>> datosxproducto) {
        List<Map<String, String>> listaDatosCarta = new ArrayList<Map<String, String>>();
        for (int i = 0; i < datosxproducto.size(); i++) {
            
            if (tipoCarta.equals("Productividad_Academica")) {
                if (IsProductividad(datosxproducto.get(i).get("TIPO_PRODUCTO"))) {
                    listaDatosCarta.add(datosxproducto.get(i));
                }
            } else if (tipoCarta.equals(datosxproducto.get(i).get("TIPO_PRODUCTO"))) {
                listaDatosCarta.add(datosxproducto.get(i));
            }
        }
        return listaDatosCarta;
    }

    private boolean IsProductividad(String tipoProducto) {
        switch (tipoProducto) {
            case "Articulo":
                return true;
            case "Produccion_de_Video_Cinematograficas_o_Fonograficas":
                return true;
            case "Libro":
                return true;
            case "Capitulo_de_Libro":
                return true;
            case "Premios_Nacionales_e_Internacionales":
                return true;
            case "Patente":
                return true;
            case "Traduccion_de_Libros":
                return true;
            case "Obra_Artistica":
                return true;
            case "Produccion_Tecnica":
                return true;
            case "Produccion_de_Software":
                return true;
            case "Ponencias_en_Eventos_Especializados":
                return true;
            case "Publicaciones_Impresas_Universitarias":
                return true;
            case "Estudios_Posdoctorales":
                return true;
            case "Reseñas_Críticas":
                return true;
            case "Traducciones":
                return true;
            case "Direccion_de_Tesis":
                return true;
            case "Evaluacion_como_par":
                return true;
            default:
                return false;
        }
    }

    private void getConsecutivoCarta(Document documento, String tipoCarta) throws DocumentException, IOException, Exception {

        if (tipoCarta.equals("Ascenso_en_el_Escalafon_Docente")) {
            p = new Paragraph(10);
            p.setAlignment(right);
            p.setFont(Fonts.SetFont(Color.black, 11,Fonts.BOLD));
            p.add("CIARP-" + "N/A" + "-" + anio + "\n\n");
            documento.add(p);
        } else {
            p = new Paragraph(10);
            p.setAlignment(right);
            p.setFont(Fonts.SetFont(Color.black, 11,Fonts.BOLD));
            p.add("CIARP-" + conseAdd + "-" + anio + "\n\n");
            documento.add(p);
        }
    }
    private void GenerarCartaIngreso(Document documento, Map<String, String> listadocentes, List<Map<String, String>> datosproductosxcartas) throws DocumentException, IOException, Exception {
        Map<String, String> respuesta = new HashMap<>();
        int banderar = 0;
        String docente = Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        String tipoProducto = Utilidades.Utilidades.decodificarElemento(listadocentes.get("TIPO_PRODUCTO"));
        String nombreProducto = Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_SOLICITUD"));

        p = new Paragraph(10);
        p.setAlignment(right);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        Chunk c = new Chunk("ASUNTO: ", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
        p.add(c);
        p.add("Respuesta a solicitud de ingreso a la carrera docente \n");
        documento.add(p);
        asunto = "Respuesta a solicitud de ingreso a la carrera docente";
 p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                p.add("Cordial saludo, \n");
                documento.add(p);
        p = new Paragraph(10);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("Cordial saludo, \n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
        try {
            p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception("" + ex.getMessage());
        }

        p.add(" estudió su solicitud de ingreso a la carrera docente \n");
        p.add(" Una vez revisada la documentación y verificado el cumplimiento de la norma el Comité determinó " + Utilidades.Utilidades.decodificarElemento(datosproductosxcartas.get(0).get("DECISION")) + " \n");
        documento.add(p);

        if (listadocentes.get("RESPUESTA_CIARP").equals("Aprobado")) {
            banderar = 1;
            
        }
        
        if (banderar == 1) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
            p.add("No obstante, se indica que el Rector de la Universidad "
                    + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                    + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                    + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
            documento.add(p);
        }
        DatosParaExcel(conseAdd,docente,asunto);
        conseAdd++;
    }

    private void GenerarCartaRevision(Document documento, Map<String, String> listadocentes, List<Map<String, String>> datosProductosxCarta) throws DocumentException, IOException, Exception {
      
        String docente=Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        p = new Paragraph(10);
        p.setAlignment(right);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        Chunk c = new Chunk("ASUNTO: ", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
        p.add(c);
        p.add("Respuesta a comunicación enviada \n");
        documento.add(p);
        asunto = "Respuesta a comunicación enviada";
 p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                p.add("Cordial saludo, \n");
                documento.add(p);
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
        try {
            p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception("" + ex.getMessage());
        }
        p.add(" estudió su comunicación en la cual manifiesta:" + " \n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setIndentationLeft(13);
        p.setIndentationRight(12);
        p.setFont(Fonts.SetFont(Color.black, 8,Fonts.ITALIC));
        try {
            String[] carta = Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("NOMBRE_SOLICITUD")).split(":");
            

            p.add(carta[1] + "\n");
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception("" + ex.getMessage());

        }
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("" + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("DECISION")) + "\n");
        documento.add(p);
        
        DatosParaExcel(conseAdd, docente, asunto);
        conseAdd++;
    }

    private void GenerarCartaAscenso(Document documento, Map<String, String> listadocentes, List<Map<String, String>> datosProductosxCarta) throws DocumentException, IOException, Exception{
     int banderar=0;  
     String docente=Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                    Chunk c = new Chunk("ASUNTO: ", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
                    p.add(c);
                    p.add("Respuesta a solicitud de ascenso en el escalafón docente. \n");
                    documento.add(p);
                    asunto = "Respuesta a solicitud de ascenso en el escalafón docente";
     p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                p.add("Cordial saludo, \n");
                documento.add(p);
                    String retroactividad = "";
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                    try {
                        p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                       throw new Exception(""+ex.getMessage());
                    }

                    p.add(" estudió la solicitud de ascenso " + (datosProductosxCarta.get(0).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " docente de planta " + datosProductosxCarta.get(0).get("NOMBRE_DOCENTE") + ", de la categoría " + datosProductosxCarta.get(0).get("CATEGORIA_DOCENTE")
                            + " a " + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("NOMBRE_SOLICITUD")) + " \n \n");
                    try {
                        p.add("Después de revisar el cumplimiento de los requisitos, el Comité decidió aprobar la promoción en el escalafón " + (datosProductosxCarta.get(0).get("SEXO").equals("M") ? "del " : "de la ") + "docente y asignarle " + getNumeroDecimal(datosProductosxCarta.get(0).get("PUNTOS")) + " (" + ValidarNumeroDec(datosProductosxCarta.get(0).get("PUNTOS")) + ") puntos salariales");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                       throw new Exception(""+ex.getMessage());
                    }
                    if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("RETROACTIVIDAD")).equals(listaDatosacta.get("FECHA_ACTA")) || Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("RETROACTIVIDAD")).equals("N/A")) {
                        
                        retroactividad += " a partir de la fecha de la presente sesión.";
                    } else if (Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("RETROACTIVIDAD")).length() > 10) {
                        retroactividad += Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("RETROACTIVIDAD")) + ".";
                    } else {
                        try {
                            retroactividad += " a partir de " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("RETROACTIVIDAD")), 0) + ".";
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception(""+ex.getMessage());
                        }
                    }
                    p.add(retroactividad + "\n");
                    documento.add(p);

                    if (datosProductosxCarta.get(0).get("RESPUESTA_CIARP").equals("Aprobado")) {
                        banderar = 1;
                        
                    }
                    
                    if (banderar == 1) {
                        p = new Paragraph(10);
                        p.setAlignment(justificado);
                        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                        p.add("No obstante, se indica que el Rector de la Universidad "
                                + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                                + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                                + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
                        documento.add(p);
                    }
                    DatosParaExcel(conseAdd, docente, asunto);
                    
    }

    private void GenerarCartaTitulacion(Document documento, Map<String, String> listadocentes, List<Map<String, String>> datosProductosxCarta) throws DocumentException, IOException, Exception {
       int banderar =0;
       String docente=Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        p = new Paragraph(10);
                    p.setAlignment(right);
                    p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                    Chunk c = new Chunk("ASUNTO: ", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
                    p.add(c);
                    p.add("Respuesta a solicitud de puntos por titulación. \n");
                    documento.add(p);
                    asunto = "Respuesta a solicitud de puntos por titulación";
                     p = new Paragraph(10);
                p.setAlignment(left);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                p.add("Cordial saludo, \n");
                documento.add(p);
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                    p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
                    try {
                        p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                       throw new Exception(""+ex.getMessage());
                    }

                    p.add(" estudió su solicitud de puntos por el título de " + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("NOMBRE_SOLICITUD")) + " \n");
                    p.add(" Una vez revisada la documentación y verificado el cumplimiento de la norma el Comité determinó " + Utilidades.Utilidades.decodificarElemento(datosProductosxCarta.get(0).get("DECISION")) + " \n");
                    documento.add(p);

                    if (datosProductosxCarta.get(0).get("RESPUESTA_CIARP").equals("Aprobado")) {
                        banderar = 1;
                        
                    }
                    
                    if (banderar == 1) {
                        p = new Paragraph(10);
                        p.setAlignment(justificado);
                        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                        p.add("No obstante, se indica que el Rector de la Universidad "
                                + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                                + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                                + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
                        documento.add(p);
                    }
                    DatosParaExcel(conseAdd, docente, asunto);
                    conseAdd++;
    }

    private void GenerarCartaProductividad(Document documento, Map<String, String> listadocentes, List<Map<String, String>> datosProductosxCarta) throws DocumentException, IOException, Exception{
        int banderar = 0;
        int bandPtsSal = 0;
        int bandPtsBon = 0;
        String docente = Utilidades.Utilidades.decodificarElemento(listadocentes.get("NOMBRE_DOCENTE"));
        
        p = new Paragraph(10);
        p.setAlignment(right);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        Chunk c = new Chunk("ASUNTO: ", Fonts.SetFont(Color.black, 11,Fonts.BOLD));
        p.add(c);
        p.add("Respuesta a solicitud de puntos por productividad académica. \n");
        documento.add(p);
        asunto = "Respuesta a solicitud de puntos por productividad académica";
        p = new Paragraph(10);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("Cordial saludo, \n");
        documento.add(p);
        
        double puntos_salariales = 0;
        double puntos_bonificacion = 0;
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
        p.add("El objetivo de la presente es comunicarle que el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión No. ");
        try {
            p.add(listaDatosacta.get("No_ACTA") + " del " + fechaEnletras(listaDatosacta.get("FECHA_ACTA"), 0));
        } catch (Exception ex) {
            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
            throw new Exception("" + ex.getMessage());
        }

        p.add(" estudió " + (datosProductosxCarta.size() > 1 ? "sus solicitudes de puntos por los productos: " : "su solicitud de puntos por el producto:") + " \n");
        documento.add(p);

        for (int i = 0; i < jerarquiaProducto.size(); i++) {
            List<Map<String, String>> listaProductosxJerarquia = data_list(3, datosProductosxCarta, new String[]{"TIPO_PRODUCTO<->" + jerarquiaProducto.get(i).get("PRODUCTO")});

            if (listaProductosxJerarquia.size() > 0) {
                
                for (int l = 0; l < listaProductosxJerarquia.size(); l++) {
                    
                    p = new Paragraph(10);
                    p.setAlignment(justificado);
                    p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                    String nameproduct = getNOMBREPRODUCTO(listaProductosxJerarquia.get(l));
                    
                    p.add("• " + listaProductosxJerarquia.get(l).get("TIPO_PRODUCTO").replace("_", " ") + ":" + nameproduct + "\n");
                    try {
                        p.add("Por este producto el Comité determinó " + Utilidades.Utilidades.decodificarElemento(getDecision(listaProductosxJerarquia.get(l))) + "\n ");
                    } catch (Exception ex) {
                        Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                        throw new Exception("" + ex.getMessage());
                    }
                    documento.add(p);
                    
                    if (listaProductosxJerarquia.get(l).get("TIPO_PUNTAJE").equals("puntos salariales")) {
                        try {
                            puntos_salariales += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get(l).get("PUNTOS").replace(",", ".")));
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception("" + ex.getMessage());
                        }
                        bandPtsSal = 1;
                    }

                    if (listaProductosxJerarquia.get(l).get("TIPO_PUNTAJE").equals("puntos de bonificacion")) {
                        try {
                            puntos_bonificacion += Double.parseDouble(ValidarNumero(listaProductosxJerarquia.get(l).get("PUNTOS").replace(",", ".")));
                        } catch (Exception ex) {
                            Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                            throw new Exception("" + ex.getMessage());
                        }
                        bandPtsBon = 1;
                    }
                    if (listaProductosxJerarquia.get(l).get("RESPUESTA_CIARP").equals("Aprobado")) {
                        banderar = 1;
                        
                    }

                }
            }

        }
        
        if (banderar == 1) {
            

            if (bandPtsSal == 1) {
                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
                try {
                    p.add("Para un total de " + getNumeroDecimal(Double.toString(puntos_salariales)) + " (" + ValidarNumeroDec("" + puntos_salariales) + ") puntos salariales por la productividad presentada. \n ");
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                    throw new Exception("" + ex.getMessage());
                }
                documento.add(p);
            }
            
            if (bandPtsBon == 1) {

                p = new Paragraph(10);
                p.setAlignment(justificado);
                p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));

                try {
                    p.add("Para un total de " + getNumeroDecimal(Double.toString(puntos_bonificacion)) + " (" + ValidarNumeroDec("" + puntos_bonificacion) + ") puntos de bonificación por la productividad presentada. \n ");
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                    throw new Exception("" + ex.getMessage());
                }
                documento.add(p);
            }
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 11,Fonts.NORMAL));
            p.add("No obstante, se indica que el Rector de la Universidad "
                    + "en caso de no estar de acuerdo con la decisión de aprobación tomada por el CIARP, podrá objetarla o rechazarla,"
                    + " en consecuencia, procederá a devolverla con una sustentación de su desacuerdo para que se revise nuevamente el caso, "
                    + "de conformidad con lo dispuesto en el Parágrafo del Artículo Noveno del Acuerdo Superior N° 021 de 2009. \n ");
            documento.add(p);
        }
        DatosParaExcel(conseAdd, docente, asunto);
        conseAdd++;
        
    }

    private void DatosParaExcel(int conseAdd, String docente, String asunto) {
        
        datos1 = new HashMap<>();
                    
                    datos1.put("CONSECUTIVO", " " + conseAdd);
                    datos1.put("ASUNTO", asunto);
                    datos1.put("DIRIGIDO_A", docente);
                    datos1.put("FECHA", fecha);
                    datos2.add(datos1);
    }
}
