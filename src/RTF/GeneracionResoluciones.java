/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package RTF;

import Excel.ControlArchivoExcel;
import Excel.gestorInformes;
import Utilidades.Fonts;
import static RTF.GeneracionActas.DatosNumeralesActa;
import static RTF.GeneracionActas.Datosacta;
import static RTF.GeneracionActas.URL;
import static RTF.GeneracionCartas.listaDatosacta;
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
import com.lowagie.text.Rectangle;
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

/**
 *
 * @author rjulio, rramos, kdelosreyes
 */
public class GeneracionResoluciones {

    static List<Map<String, String>> normaProducto = new ArrayList<>();
    static Map<String, String> listaProyecto = new HashMap<>();
    public String ruta;
    public String semestre;
    static String URL;
    static List<Map<String, String>> listaDatos = new ArrayList<>();
    static Map<String, String> DatosResoluciones = new HashMap<>();
    static ArrayList<ArrayList<Map<String, String>>> DatosNumeralesResoluciones = new ArrayList<ArrayList<Map<String, String>>>();
    ArrayList<Map<String, String>> datos2 = new ArrayList<>();
    Map<String, String> respuesta = new HashMap<>();
    static int bans = 0;
    static String lineaDeTexto;
    BaseFont calibriFont;
    BaseFont arialFont;

    public GeneracionResoluciones() {
        normaProducto = new ArrayList<>();
        listaProyecto = new HashMap<>();
        ruta = "";
        semestre = "";
        URL = "";
        listaDatos = new ArrayList<>();
        DatosResoluciones = new HashMap<>();
        DatosNumeralesResoluciones = new ArrayList<ArrayList<Map<String, String>>>();
        datos2 = new ArrayList<>();
        bans = 0;
        lineaDeTexto = "";
         
        InicializarNorma();
        
    }

    public void InicializarNorma() {
        Map<String, String> auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ingreso_a_la_Carrera_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ingreso a la Carrera Docente");
        auxiliarJerarquia.put("NORMA", "Artículo 14 del Acuerdo Superior N° 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ascenso_en_el_Escalafon_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ascenso en el Escalafón Docente");
        auxiliarJerarquia.put("NORMA", "Artículo 27 del Acuerdo Superior N° 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Titulacion");
        auxiliarJerarquia.put("NPRODUCTO", "Titulación");
        auxiliarJerarquia.put("NORMA", "Artículo 7 del Decreto 1279 del 2002, Artículo Primero del Acuerdo 001 de 2004 del Grupo de Seguimiento al Decreto 1279 de 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Articulo");
        auxiliarJerarquia.put("NPRODUCTO", "Artículo");
        auxiliarJerarquia.put("NORMA", "literal a. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Video_Cinematograficas_o_Fonograficas");
        auxiliarJerarquia.put("NPRODUCTO", "Producción de Video Cinematografica o Fonografica");
        auxiliarJerarquia.put("NORMA", "literal b. numeral I, Artículo 10; literal a. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Libro");
        auxiliarJerarquia.put("NORMA", "literales c, d, e, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Capitulo_de_Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Capítulo de Libro");
        auxiliarJerarquia.put("NORMA", "literales c, d, e, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Premios_Nacionales_e_Internacionales");
        auxiliarJerarquia.put("NPRODUCTO", "Premio Nacional o Internacional");
        auxiliarJerarquia.put("NORMA", "literal f. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Patente");
        auxiliarJerarquia.put("NPRODUCTO", "Patente");
        auxiliarJerarquia.put("NORMA", "literal g. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traduccion_de_Libros");
        auxiliarJerarquia.put("NPRODUCTO", "Traducción de Libro");
        auxiliarJerarquia.put("NORMA", "literal h. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Obra_Artistica");
        auxiliarJerarquia.put("NPRODUCTO", "Obra Artistica");
        auxiliarJerarquia.put("NORMA", "literal i. numeral I, Artículo 10, literal g. numeral II, Articulo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_Tecnica");
        auxiliarJerarquia.put("NPRODUCTO", "Producción Técnica");
        auxiliarJerarquia.put("NORMA", "literal j. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Software");
        auxiliarJerarquia.put("NPRODUCTO", "Producción de Software");
        auxiliarJerarquia.put("NORMA", "literal k. numeral I, Artículo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ponencias_en_Eventos_Especializados");
        auxiliarJerarquia.put("NPRODUCTO", "Ponencia en Evento Especializado");
        auxiliarJerarquia.put("NORMA", "literal b. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Publicaciones_Impresas_Universitarias");
        auxiliarJerarquia.put("NPRODUCTO", "Publicación Impresa Universitaria");
        auxiliarJerarquia.put("NORMA", "literal c. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Estudios_Posdoctorales");
        auxiliarJerarquia.put("NPRODUCTO", "Estudio Posdoctoral");
        auxiliarJerarquia.put("NORMA", "literal d. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Reseñas_Críticas");
        auxiliarJerarquia.put("NPRODUCTO", "Reseña Crítica");
        auxiliarJerarquia.put("NORMA", "literal e. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traducciones");
        auxiliarJerarquia.put("NPRODUCTO", "Traducción");
        auxiliarJerarquia.put("NORMA", ""
                + "literal f. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Direccion_de_Tesis");
        auxiliarJerarquia.put("NPRODUCTO", "Dirección de Tesis");
        auxiliarJerarquia.put("NORMA", "literal h. numeral II, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Evaluacion_como_par");
        auxiliarJerarquia.put("NPRODUCTO", "Evaluación como par");
        auxiliarJerarquia.put("NORMA", "literal i. numeral I, Artículo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

    }

    public String Encode() {
        String cifrado = "" + System.currentTimeMillis();
        return cifrado;
    }

    public void leerArchivo() {
        ControlArchivo contArchivo = new ControlArchivo(ruta);
        contArchivo.LeerArchivo();
        BufferedReader br = contArchivo.getBuferDeLectura();

        String[] keys = new String[]{};
        int nkey = 0;
        int cont = 0;

    }

    public Map<String, String> CrearResoluciones(String ruta, String semestre) throws IOException, DocumentException {
        this.ruta = ruta;
        this.semestre = semestre;
        ControlArchivoExcel con = new ControlArchivoExcel();
         
         //<editor-fold defaultstate="collapsed" desc="Lectura Puntos Todos">
        String extP = ruta.substring(ruta.lastIndexOf(".") + 1);
        if (extP.equals("xlsx")) {
            listaDatos = con.LeerExcelDesdeAct(ruta, 2, "PUNTOS TODOS");
            listaProyecto = con.LeerExcelParametrosAct(ruta, 1, 1,"PUNTOS TODOS");
        } else {
            listaDatos = con.LeerExcelDesde(ruta, 2, "PUNTOS TODOS");
            listaProyecto = con.LeerExcelParametros(ruta, 1, 1); 
        }
        
//</editor-fold>

        
        GenerarResolucion(); 

        return respuesta;
    }

    public Map <String, String> GenerarResolucion() throws IOException, DocumentException {
        try{
            arialFont = BaseFont.createFont("C:\\windows\\Fonts\\ARIAL.TTF", "Cp1252", true);
        
        calibriFont = BaseFont.createFont("C:\\windows\\Fonts\\CALIBRI.TTF", "Cp1252", true);
        List<Map<String, String>> listaTipoResolucion = data_list(1, listaDatos, new String[]{"TIPO_RESOLUCION"});
        List<Map<String, String>> listaCargoAcademicoAdministrativo = data_list(3, listaDatos, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin"});
        List<Map<String, String>> listaCargoAcademicoAdministrativoxtipoProducto = data_list(1, listaCargoAcademicoAdministrativo, new String[]{"TIPO_PRODUCTO"});
        List<Map<String, String>> listaCargoAcademicoAdminSalarial = getTipoProductoSalarial(listaCargoAcademicoAdministrativoxtipoProducto);
        List<Map<String, String>> listaDocentesTipoResolucion = new ArrayList<>();
        
            
        int vvv = 0;
        for(Map<String, String> map : listaTipoResolucion){
            vvv++;
            
        }
        
            
        
        DatosResoluciones.put("NOMBRE_ARCHIVO", "numerales_resoluciones_");
        DatosResoluciones.put("NUMACTA", ""+semestre);
        
        //<editor-fold defaultstate="collapsed" desc="Tipo de resoluciones Cargo Acad Admin por tipoProducto sin Resolucion">
            List<Map<String, String>> ListaSinTipoResol = getSinTipoResolucion(listaTipoResolucion, listaCargoAcademicoAdministrativoxtipoProducto);
            
            if(ListaSinTipoResol.size()>0){
                for(Map<String, String> lst:ListaSinTipoResol){
                    String tipoRel = getResolucionxProducto(lst.get("TIPO_PRODUCTO"));
                    
                    //<editor-fold defaultstate="collapsed" desc="Listado Docentes x Tipos de Resolucion">
                    if (tipoRel.toUpperCase().equals("CONVALIDACION")) {
                        listaDocentesTipoResolucion = new ArrayList<>();
                        listaDocentesTipoResolucion = data_list(1, listaCargoAcademicoAdministrativo, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Convalidacion"});
                        
                    } else if (tipoRel.toUpperCase().equals("TITULACION")) { 
                        listaDocentesTipoResolucion = new ArrayList<>();
                        listaDocentesTipoResolucion = data_list(1, listaCargoAcademicoAdministrativo, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Titulacion"});
                        
                    } else if (tipoRel.toUpperCase().equals("Ascenso_en_el_escalafon".toUpperCase())) {
                        listaDocentesTipoResolucion = new ArrayList<>();
                        listaDocentesTipoResolucion = data_list(1, listaCargoAcademicoAdministrativo, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Ascenso_en_el_Escalafon_Docente"});
                        
                    } else if (tipoRel.toUpperCase().equals("SALARIAL")) {
                        
                        List<Map<String, String>> listaDatosAcad = new ArrayList<>();
                        List<Map<String, String>> listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
        
                        for (Map<String, String> datos : listaCargoAcademicoAdminSalarial) {
                            listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
                            listaDocentesTipoResolucionAuxiliar = data_list(1, listaCargoAcademicoAdministrativo, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->" + datos.get("TIPO_PRODUCTO")});

                            if (listaDocentesTipoResolucionAuxiliar.size() > 0) {
                                listaDocentesTipoResolucion.addAll(listaDocentesTipoResolucionAuxiliar);
                            }
                        }

                    }
//</editor-fold>
                    
                    for (int j = 0; j < listaDocentesTipoResolucion.size(); j++) {
                        URL = "C:\\CIARP\\RESOLUCIONES\\" + tipoRel + "\\" + listaDocentesTipoResolucion.get(j).get("NOMBRE_DEL_DOCENTE") + ".rtf";
                        
                        Document documento = new Document(new Rectangle(612, 936));
                        documento.setMargins(70, 70, 69, 55);
                        RtfWriter2.getInstance(documento, new FileOutputStream(URL));

                        documento.open();
        
                        if (tipoRel.equals("Salarial")) {
                            GenerarResolSalarial(documento, listaDocentesTipoResolucion.get(j).get("CEDULA"), tipoRel);
                        }else if(tipoRel.equals("Ascenso_en_el_escalafon")){
                            GenerarResolAscenso(documento, listaDocentesTipoResolucion.get(j)); 
                        }else if(tipoRel.equals("Bonificacion")){
                            GenerarResolBonificacion(documento, listaDocentesTipoResolucion.get(j).get("CEDULA"), tipoRel);
                        }else if(tipoRel.equalsIgnoreCase("Convalidacion")){
                            GenerarResolConvalidacion(documento, listaDocentesTipoResolucion.get(j));
                        }else if(tipoRel.equals("Ingreso_carrera_docente")){
                            GenerarResolIngreso(documento, listaDocentesTipoResolucion.get(j));
                        }else if(tipoRel.equals("Titulacion")){
                            GenerarResolTitulacion(documento, listaDocentesTipoResolucion.get(j));
                        }

                        documento.close();
                        File f = new File(URL);
                        f.createNewFile();
                        
                    }
                    
                    
                }
            }
            
        //</editor-fold>
        
        
        // # DOCENTES:   10   ,      2          6         
        for (int i = 0; i < listaTipoResolucion.size(); i++) {//ARTICULO, TITULACION, PONENCIA
            
            if (listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Cargo_acad_admin")) {
                continue;
            }

            //<editor-fold defaultstate="collapsed" desc="CONDICION INICIAL">
            if (bans == 1 && listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL")) {
                continue;
            } else if (bans == 1 && listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL_COLCIENCIAS")) {
                continue;
            }
            //</editor-fold>

            listaDocentesTipoResolucion = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->" + listaTipoResolucion.get(i).get("TIPO_RESOLUCION")});

            //<editor-fold defaultstate="collapsed" desc="COLSALARIAL">
            List<Map<String, String>> listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
            if (bans == 0 && listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL")) {// ADICIONAR SALARIAL_COLCIENCIA
                listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Salarial_colciencias"});

                listaDocentesTipoResolucionAuxiliar.addAll(listaDocentesTipoResolucion);

                listaDocentesTipoResolucion = new ArrayList<>();
                listaDocentesTipoResolucion = data_list(1, listaDocentesTipoResolucionAuxiliar, new String[]{"CEDULA"});
                bans = 1;
            } else if (bans == 0 && listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL_COLCIENCIAS")) {// ADICIONAR SALARIAL
                listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Salarial"});
                listaDocentesTipoResolucionAuxiliar.addAll(listaDocentesTipoResolucion);

                listaDocentesTipoResolucion = new ArrayList<>();

                listaDocentesTipoResolucion = data_list(1, listaDocentesTipoResolucionAuxiliar, new String[]{"CEDULA"});
                bans = 1;
            }
            //</editor-fold>
            
            

            ///////////////CARGO ACADEMICO ADMINISTRATIVO///////////////
            if (listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("CONVALIDACION")) {
                listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
                listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Convalidacion"});
                if (listaDocentesTipoResolucionAuxiliar.size() > 0) {
                    listaDocentesTipoResolucionAuxiliar.addAll(listaDocentesTipoResolucion);

                    listaDocentesTipoResolucion = new ArrayList<>();
                    listaDocentesTipoResolucion = data_list(1, listaDocentesTipoResolucionAuxiliar, new String[]{"CEDULA"});
                }
            } else if (listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("TITULACION")) {
                listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
                listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Titulacion"});
                if (listaDocentesTipoResolucionAuxiliar.size() > 0) {
                    listaDocentesTipoResolucionAuxiliar.addAll(listaDocentesTipoResolucion);
                    listaDocentesTipoResolucion = new ArrayList<>();
                    listaDocentesTipoResolucion = data_list(1, listaDocentesTipoResolucionAuxiliar, new String[]{"CEDULA"});
                }
            } else if (listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("Ascenso_en_el_escalafon".toUpperCase())) {
                listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
                listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Ascenso_en_el_Escalafon_Docente"});
                if (listaDocentesTipoResolucionAuxiliar.size() > 0) {
                    listaDocentesTipoResolucionAuxiliar.addAll(listaDocentesTipoResolucion);

                    listaDocentesTipoResolucion = new ArrayList<>();
                    listaDocentesTipoResolucion = data_list(1, listaDocentesTipoResolucionAuxiliar, new String[]{"CEDULA"});
                }

            } else if ((listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL_COLCIENCIAS")
                    || listaTipoResolucion.get(i).get("TIPO_RESOLUCION").toUpperCase().equals("SALARIAL"))) {
                
                List<Map<String, String>> listaDatosAcad = new ArrayList<>();

                for (Map<String, String> datos : listaCargoAcademicoAdminSalarial) {
                    listaDocentesTipoResolucionAuxiliar = new ArrayList<>();
                    listaDocentesTipoResolucionAuxiliar = data_list(1, listaDatos, new String[]{"CEDULA"}, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->" + datos.get("TIPO_PRODUCTO")});

                    if (listaDocentesTipoResolucionAuxiliar.size() > 0) {
                        listaDatosAcad.addAll(listaDocentesTipoResolucionAuxiliar);
                    }
                }
               
                if (listaDatosAcad.size() > 0) {
                    listaDatosAcad.addAll(listaDocentesTipoResolucion);

                    listaDocentesTipoResolucion = new ArrayList<>();
                    listaDocentesTipoResolucion = data_list(1, listaDatosAcad, new String[]{"CEDULA"});
                }
            }
            
            

            for (int j = 0; j < listaDocentesTipoResolucion.size(); j++) {
                URL = "C:\\CIARP\\RESOLUCIONES\\" + listaTipoResolucion.get(i).get("TIPO_RESOLUCION") + "\\" + Utilidades.Utilidades.decodificarElemento(listaDocentesTipoResolucion.get(j).get("NOMBRE_DEL_DOCENTE")) + ".rtf";
                
                Document documento = new Document(new Rectangle(612, 936));
                documento.setMargins(70, 70, 69, 55);
                RtfWriter2.getInstance(documento, new FileOutputStream(URL));

                documento.open();

                if (listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Salarial") || listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Salarial_colciencias")) {
                    GenerarResolSalarial(documento, listaDocentesTipoResolucion.get(j).get("CEDULA"), listaTipoResolucion.get(i).get("TIPO_RESOLUCION"));
                }else if(listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Ascenso_en_el_escalafon")){
                    GenerarResolAscenso(documento, listaDocentesTipoResolucion.get(j)); 
                }else if(listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Bonificacion")){
                    GenerarResolBonificacion(documento, listaDocentesTipoResolucion.get(j).get("CEDULA"), listaTipoResolucion.get(i).get("TIPO_RESOLUCION"));
                }else if(listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equalsIgnoreCase("Convalidacion")){
                    GenerarResolConvalidacion(documento, listaDocentesTipoResolucion.get(j));
                }else if(listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Ingreso_carrera_docente")){
                    
                    GenerarResolIngreso(documento, listaDocentesTipoResolucion.get(j));
                }else if(listaTipoResolucion.get(i).get("TIPO_RESOLUCION").equals("Titulacion")){
                    GenerarResolTitulacion(documento, listaDocentesTipoResolucion.get(j));
                }
            
                
                documento.close();
                File f = new File(URL);
                f.createNewFile();
                
            }
        }
        DatosNumeralesResoluciones.add(datos2);
       
        gestorInformes gi = new gestorInformes(DatosResoluciones, DatosNumeralesResoluciones);
        gi.iniciar();
        respuesta.put("ESTADO", "OK");
        respuesta.put("MENSAJE", "Las resoluciones se generaron correctamente.");

        
        }catch(Exception e){
            e.printStackTrace();
            System.out.println("erro->"+e.getMessage());
            respuesta.put("ESTADO", "ERROR");
            respuesta.put("MENSAJE", e.getMessage());
            
        }
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

    private Map<String, String> GenerarResolSalarial(Document documento, String identificacion, String tipo_resolucion) throws DocumentException, IOException {
        try{
        String NOM_DOCENTE = "";
        
        List<Map<String, String>> listaDatosDocentes = data_list(3, listaDatos, new String[]{"TIPO_RESOLUCION<->Salarial", "CEDULA<->" + identificacion});

        List<Map<String, String>> listaDatosDocentesAux = data_list(3, listaDatos, new String[]{"TIPO_RESOLUCION<->Salarial_colciencias", "CEDULA<->" + identificacion});

        List<Map<String, String>> listaDatosDocentesCargoAcad = data_list(3, listaDatos, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "CEDULA<->" + identificacion});
        Double totalpuntos = 0.0;
        if (listaDatosDocentesAux.size() > 0) {
            listaDatosDocentes.addAll(listaDatosDocentesAux);
        }
        if (listaDatosDocentesCargoAcad.size() > 0) {
            listaDatosDocentes.addAll(listaDatosDocentesCargoAcad);
        }
        Map<String, String> datos1 = new HashMap<>();
        
        
        
        List<Map<String, String>> listaActas = data_list(1, listaDatosDocentes, new String[]{"ACTA"});

        NOM_DOCENTE = Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));

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
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");
     

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

        //<editor-fold defaultstate="collapsed" desc="HEADER">
        
        Table headerTable;
        Table headerTableTxt;


        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
         Image imgM = Image.getInstance("C:\\CIARP\\under.png");

        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "“Por la cual se autoriza el reconocimiento y pago de puntos salariales "
                + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " "
                + NOM_DOCENTE + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>

        celda = new Cell(new Paragraph(resolucion+ "”", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
       
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);

            //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(right);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
     
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);
            //</editor-fold>

        
        RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
      
        //</editor-fold>

        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(fh3);
        p.setAlignment(centrado);
        c = new Chunk("CONSIDERANDO:\n", af10b);
        p.add(c);
        documento.add(p);

        if(listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE TIEMPO COMPLETO") || 
            listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE MEDIO TIEMPO")    ){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que el Artículo 70 del Acuerdo Superior N° 007 de 2003 establece las condiciones para la valoración de la productividad académica de los docentes ocasionales. \n");
            documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setFont(af10);
        p.setAlignment(justificado);
        p.add("Que el Capítulo III del Decreto N° 1279 de 2002, establece los factores y criterios a tener en cuenta para la modificación de puntos salariales de los docentes amparados por dicho régimen, siendo la productividad académica, uno de los factores incidentes en este proceso, según lo establecido en el Literal C., del Artículo 12 y el Artículo 16 de la disposición en cita.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(af10);
        p.setAlignment(justificado);
        p.add("Que el Comité Interno de Asignación y Reconocimiento de Puntaje – CIARP, ha considerado para la asignación de puntaje de los docentes de planta, los criterios de evaluación y asignación que para el efecto ha establecido el Grupo de Seguimiento del Régimen Salarial y Prestacional de los Profesores Universitarios mediante el Acuerdo N° 001 de 04 de marzo de 2004.\n");
        documento.add(p);
        
        if(listaDatosDocentesCargoAcad.size()>0){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Artículo 17 del referido decreto, establece los criterios a tener en cuenta para la modificación de salario "+
                    "de los docentes que realizan actividades académico-administrativas, disponiéndose, en igual sentido, en el Artículo 62 de la mencionada "+
                    "disposición, que el Grupo de Seguimiento al régimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la información a nivel nacional, y además adecuar los criterios "+
                    "y efectuar los ajustes a las metodologías de evaluación aplicadas por los Comités Internos de Asignación de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al Régimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N° 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad académica para los docentes que asuman cargos académico-administrativos, este "+
                    "señaló que independientemente de haber elegido entre la remuneración del cargo que va a desempeñar y la que le corresponde como docente, "+
                    "solo se podrá hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades académico-administrativas, de conformidad con el Artículo 17 del Decreto 1279 de 2002 y el parágrafo 1 del Artículo 60 del Acuerdo N° 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que mediante Resolución Rectoral Nº "+
                    listaDatosDocentes.get(0).get("RESOL_ENCARGO")+" "+
                    (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get(0).get("TIPO_VINCULACION"));
            c = new Chunk(" "+Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE)+", ",af10b);
            p.add(c);
                    
            p.add("fue "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "comisionado" : "comisionada")+" para ejercer un cargo de libre nombramiento y remoción dentro de la Institución como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tomó posesión mediante Acta N° "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE), af10b);
        p.add(c);
        p.add(", presentó solicitud de asignación de puntos salariales por" + (listaDatosDocentes.size()==1?" un":numeroEnLetras(listaDatosDocentes.size())) + " (" + listaDatosDocentes.size() + ")");
        
        
        String resultif="";
        if (listaDatosDocentes.get(0).get("TIPO_PRODUCTO").equals("Articulo")||listaDatosDocentes.get(0).get("TIPO_PRODUCTO").equals("Patente")||listaDatosDocentes.get(0).get("TIPO_PRODUCTO").equals("Art_Col")){
            if(listaDatosDocentes.size()>1){
                resultif=" productos clasificados y categorizados ";
            }else{
                resultif=" producto clasificado y categorizado ";
            }
        }else{
           if(listaDatosDocentes.size()>1){
                resultif=" productos revisados por pares evaluadores, clasificados y categorizados ";
            }else{
                resultif=" producto revisado por pares evaluadores, clasificado y categorizado ";
            }
        }
        p.add(resultif);

        for (int i = 0; i < listaActas.size(); i++) {
            if (i > 0 && i < listaActas.size() - 1) {
                p.add(", ");
            } else if (i> 0 && i == listaActas.size() - 1) {
                p.add(" y ");
            }
            
            List<Map<String, String>> listadatosxActasxnumerales = data_list(1, listaDatosDocentes, new String[]{"NUMERAL_ACTA_CIARP"}, new String[]{"ACTA<->" + listaActas.get(i).get("ACTA")});
            for(int j =0 ; j < listadatosxActasxnumerales.size(); j++){
                if (j > 0 && j < listadatosxActasxnumerales.size() - 1) {
                    p.add(", ");
                } else if (j> 0 && j == listadatosxActasxnumerales.size() - 1) {
                    p.add(" y ");
                }
                p.add("en el ítem "+listadatosxActasxnumerales.get(j).get("NUMERAL_ACTA_CIARP"));
            }
            try{
            p.add(" del Acta N° " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            
        }

        p.add(".\n");
        documento.add(p);

        if (listaDatosDocentesAux.size() > 0) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            if(listaDatosDocentesAux.get(0).get("RETROACTIVIDAD").length()==10){
                try{
            p.add("Que la base de datos de COLCIENCIAS para Revistas Internacionales fue actualizada el día " + fechaEnletras(listaDatosDocentesAux.get(0).get("RETROACTIVIDAD"),0) + ", según la página web www.colciencias.gov.co.\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("TIPO_PRODUCTO")));
                                 if(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
                }
            documento.add(p);
        }

        String norma = "";
        String numerales= "";
        Double totalp = 0.0;
        for (int j = 0; j < listaActas.size(); j++) {
            
            List<Map<String, String>> listadatosxActas = data_list(3, listaDatosDocentes, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA")});
            
            List<Map<String, String>> listadatosxActasxTp = data_list(1, listaDatosDocentes, new String[]{"TIPO_PRODUCTO"}, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA")});
            norma = "";
            numerales = "";
            totalp = getSumaPuntos(listadatosxActas);
            totalpuntos += totalp;
            for (int h = 0; h < listadatosxActasxTp.size(); h++) {    
                List<Map<String, String>> listadatosxActasTP = data_list(3, listaDatosDocentes, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA"), "NUMERAL_ACTA_CIARP<->" + listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP")});
                ////NORMAS PRIMER Y SEGUNDO PRODUCTO del numeral 
                if (!Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(h).get("NORMA")).equals("#N/D")  && !Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(h).get("NORMA")).equals("#N/A")
                        && !listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ARTICULO")
                        && !listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ART_COL")) {
                    numerales += (numerales.equals("")?"":", ")+"ítem "+ 
                            getPosicionesNumeral(listadatosxActasTP)+
                      
                            " del numeral "+listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP");
                    norma += (norma.equals("")?"":", ")+Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(h).get("NORMA"));
                } else {
                    if (!listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ARTICULO")
                            && !listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ART_COL")) {
                        numerales += (numerales.equals("")?"":", ")+"Ítem "+ 
                                getPosicionesNumeral(listadatosxActasTP)+
                               
                                " del numeral "+listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP");    
                        norma += (norma.equals("")?"":", ")+getNormaProducto(listadatosxActasxTp.get(h).get("TIPO_PRODUCTO"));
                    } else {//ARTICULOS
                        
                        List<Map<String, String>> listaxcategoria = data_list(10, listadatosxActas, new String[]{"TIPO_PRODUCTO", "CATEGORIA"}, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA")});
                        String norm = "";
                        String num = "";
                        
                        for (int hh = 0; hh < listaxcategoria.size(); hh++) {
                            
                            if(hh > 0 && hh < listaxcategoria.size()-1){
                        
                                norm +=", ";
                            }else if(hh> 0 && hh == listaxcategoria.size()-1){
                               
                                norm +=" y " +(norma.equals("")?"":", ")+getNormaProducto(listadatosxActasxTp.get(h).get("TIPO_PRODUCTO"));;
                            }
                            
                            

                            
                            if(listaxcategoria.get(hh).get("CATEGORIA").equals("A1")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "ítem B, literal a. A.1";
                                }else{
                                        norm += "ítem A, literal a. A.1";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("A2")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "ítem B, literal a. A.2";
                                }else{
                                        norm += "ítem A, literal a. A.2";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("B")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "ítem B, literal a. A.3";
                                }else{
                                        norm += "ítem A, literal a. A.3";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("C")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "ítem B, literal a. A.4";
                                }else{
                                        norm += "ítem A, literal a. A.4";
                                        }
                            }
                        }
                        num+= "Ítem "+ getPosicionesNumeral(listadatosxActasTP);
                        num += " del numeral "+listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP");
                        
                        norm += ", artículo 10 del Decreto 1279";
                        norma += (norma.equals("")?"":", ")+norm;
                        numerales += (numerales.equals("")?"":", ")+num;    
                    }
                }
            }
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);

            String numeroletras = getNumeroDecimal(""+listadatosxActas.size());
            
            try{
            p.add("Que en virtud de lo anterior y conforme a lo señalado en "
                + norma
            
                +" el Comité Interno de Asignación y Reconocimiento de Puntaje en sesión realizada el día "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+", "
                +"mediante Acta N° "+listaActas.get(j).get("ACTA")+" decidió asignar "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al docente " : "a la docente "));
           } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            c = new Chunk(NOM_DOCENTE+ " ",af10b);
            p.add(c);
            try{
            p.add(
                getNumeroDecimal(""+totalp)+" ("+ValidarNumeroDec(""+totalp)+")"
               
                +" puntos salariales, "
                +"distribuidos en " +(listadatosxActas.size()>1 ? "los " + numeroletras:((listadatosxActas.size()== 1?numeroletras.substring(0, 2):numeroletras)))
                    
                    +" ("+listadatosxActas.size()+") "+(listadatosxActas.size()>1 ? "productos":"producto")+" según el orden establecido en el aparte de “Decisión” del "
                +numerales +" del Acta N° "+listaActas.get(j).get("ACTA")+" de "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+".\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            documento.add(p);
            norma="";
        }
            
        p = new Paragraph(10);
        p.setFont(af10);
        p.setAlignment(justificado);
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA, se determina dos (2) veces al año el total de puntos que corresponde a cada docente, conforme lo señala el Artículo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);
        p = new Paragraph(10);
        p.setFont(af10);
        p.setAlignment(justificado);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(fh3);
        p.setAlignment(centrado);
        c = new Chunk("RESUELVE:\n", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Reconocer y pagar "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ")+listaDatosDocentes.get(0).get("TIPO_VINCULACION")+" ");
        c = new Chunk("" + NOM_DOCENTE + ",", af10b);
        p.add(c);
       
        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? " identificado con " : " identificada con "));
        if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk("" + getNumeroDecimal(("" + totalpuntos).replace(".", ",")) + " (" + ValidarNumeroDec("" + totalpuntos) + ") puntos salariales, ", af10b);
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(0).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
        p.add((listaDatosDocentes.size()>1 ? "por los productos relacionados en la siguiente tabla: ":"por el producto relacionado en la siguiente tabla: "));
        p.add("\n");
        documento.add(p);

            ///TABLA
        //<editor-fold defaultstate="collapsed" desc="TABLA">
        int cols = (listaDatosDocentesCargoAcad.size() > 0 ? 5 : 6);
        
        float[] tamy = new float[]{5f, 13f, 50f, 8f, 13f, 20f};
        if (listaDatosDocentesCargoAcad.size() > 0) {
            tamy = new float[]{8f, 15f, 55f, 10f, 15f};
        }

        Table TableProductos = new Table(cols);
        TableProductos.setWidths(tamy);
        TableProductos.setWidth(100);

        Cell celdaProductos = new Cell(new Paragraph("N°", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Producto", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);
        celdaProductos = new Cell(new Paragraph("Nombre del Producto", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("N° Acta", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Puntos Reconocidos", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        if (listaDatosDocentesCargoAcad.size() == 0) {
            celdaProductos = new Cell(new Paragraph("Fecha a partir de la cual surten efectos fiscales", af7b));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
        }

        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            celdaProductos = new Cell(new Paragraph("" + (i + 1), af7b));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("TIPO_PRODUCTO")).equals("Art_Col") ? "Artículo" : listaDatosDocentes.get(i).get("TIPO_PRODUCTO").replace("_", " ")), af7b));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            String nombresolicitud = "";
            if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("SUBTIPO_PRODUCTO")).equals("N/A")) {
                nombresolicitud = "" + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("SUBTIPO_PRODUCTO")).replace("_", " ") + ":" + getNOMBREPRODUCTO(listaDatosDocentes.get(i));
            } else {
                nombresolicitud = getNOMBREPRODUCTO(listaDatosDocentes.get(i));
            }
            celdaProductos = new Cell(new Paragraph("" + nombresolicitud, af7));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + listaDatosDocentes.get(i).get("ACTA"), af7));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            
            try{
            celdaProductos = new Cell(new Paragraph("" + getNumeroDecimal(listaDatosDocentes.get(i).get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get(i).get("PUNTOS")) + ") puntos", af7b));
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            if (listaDatosDocentesCargoAcad.size() == 0) {
                if(listaDatosDocentes.get(i).get("RETROACTIVIDAD").length()==10){
                    try{
                    celdaProductos = new Cell(new Paragraph("" + fechaEnletras(listaDatosDocentes.get(i).get("RETROACTIVIDAD"),0), af7));
                } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
                    }else{
                    celdaProductos = new Cell(new Paragraph("" , af7));
                }
                celdaProductos.setBorder(15);
                celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
                celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
                TableProductos.addCell(celdaProductos);
            }
        }

        documento.add(TableProductos);
            //</editor-fold>

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("\n");
        c = new Chunk("Parágrafo Único: ", af10b);
        p.add(c);
        p.add("Los puntos salariales asignados en el recuadro anterior, se establecen según lo dispuesto en el "
                + "aparte de “Decisión” ");
        for (int i = 0; i < listaActas.size(); i++) {
            if (i > 0 && i < listaActas.size() - 1) {
                p.add(", ");
            } else if (i> 0 && i == listaActas.size() - 1) {
                p.add(" y ");
            }
            
            List<Map<String, String>> listadatosxActasxnumerales = data_list(1, listaDatosDocentes, new String[]{"NUMERAL_ACTA_CIARP"}, new String[]{"ACTA<->" + listaActas.get(i).get("ACTA")});
            for(int j =0 ; j < listadatosxActasxnumerales.size(); j++){
                if (j > 0 && j < listadatosxActasxnumerales.size() - 1) {
                    p.add(", ");
                } else if (j> 0 && j == listadatosxActasxnumerales.size() - 1) {
                    p.add(" y ");
                }
                p.add("en el numeral "+listadatosxActasxnumerales.get(j).get("NUMERAL_ACTA_CIARP"));
            }
            try{
            c = new Chunk(" del Acta N° " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", af10);
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            p.add(c);
        }

        p.add(" del Comité Interno de Asignación y Reconocimiento de Puntaje respectivamente. ");
        if(listaDatosDocentesCargoAcad.size()>0)
            p.add("Así mismo, surtirán efectos fiscales una vez "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el " : "la ")+"docente finalice la comisión para desempeñar el cargo de libre nombramiento y remoción que le ha sido otorgado.");
        else    
            p.add("Así mismo, tendrán efectos fiscales a partir de la fecha en que el Comité expidió "
                + "el acto formal de reconocimiento, según consta en la relación anterior.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión "
                + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk("" + Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto, envíesele la correspondiente comunicación al correo electrónico de notificación"
                + ", haciéndole saber que contra la misma procede recurso de reposición ante el mismo funcionario "
                + "que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días siguientes "
                + "a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de La Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        p.add("Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales, "
                + "a fin de que procedan a realizar el trámite correspondiente. Envíese copia de esta resolución al Comité Interno "
                + "de Asignación y Reconocimiento de Puntaje y a la hoja de vida ");

        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk("" + Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE) + ".", af10b);
        p.add(c);
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO: ", af10b);
        p.add(c);
        p.add("La presente resolución rige a partir del término de su ejecutoria.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(fh3);
        p.setAlignment(centrado);
        c = new Chunk("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n", af10b);
        p.add(c);

        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Dada en la ciudad de Santa Marta, D. T. C. H., a los");
        p.add("\n \n \n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("PABLO VERA SALAZAR");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10);
        p.add("Rector ");
        p.add("\n\n");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add("" + listaProyecto.get("PROYECTO") + "____");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add("" + listaProyecto.get("REVISO") + "____");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add("" + listaProyecto.get("REVISO_JEFE") + "____\n");
        documento.add(p);
        }catch(Exception e){
            e.printStackTrace();
        }
     return respuesta;
    }
    
    private Map<String, String> GenerarResolAscenso(Document documento, Map<String, String> listaDatosDocentes) throws BadElementException, IOException, DocumentException {
        
        int band = 0;
        Double totalPuntos = 0.0;
        Double totalPuntosxActas = 0.0;
        String item = "";
        Map<String, String> datos1 = new HashMap<>();
        int cantidaddeProductos = listaDatosDocentes.size();
        String[] numeralActas = {};

        //<editor-fold defaultstate="collapsed" desc="estilo">
        Font fh1 = new Font(arialFont);
        fh1.setSize(11);
        fh1.setStyle("bold");
        fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("bold");

        Font fh3 = new Font(arialFont);
        fh3.setSize(8);
        fh3.setColor(Color.BLACK);
        
        Font fh3c = new Font(arialFont);
        fh3c.setSize(8);
        fh3c.setColor(Color.BLACK);
        fh3c.setStyle("italic");
        
        Font fh3b = new Font(arialFont);
        fh3b.setSize(8);
        fh3b.setColor(Color.BLACK);
        fh3b.setStyle("bold");

        Font fh4 = new Font(arialFont);
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");

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
        int derecha = Paragraph.ALIGN_RIGHT;
//</editor-fold>
        
                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">

        //<editor-fold defaultstate="collapsed" desc="HEADER">
        
        
        
        RtfHeaderFooterGroup headerDif= new RtfHeaderFooterGroup();
        Table headerTable;
        Table headerTableTxt;

//      
        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
         Image imgM = Image.getInstance("C:\\CIARP\\under.png");
         
        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        String resolucion = "\"Por la cual se promociona en el Escalafón Docente a la categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))
                            + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                            + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        celda = new Cell(new Paragraph(resolucion+"", fh4));
        celda.setBorder(0);

        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
       
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);
        //</editor-fold>

        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(derecha);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
       
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        //</editor-fold>

        RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="Cuerpo Resolucion">
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el Artículo 20 del Acuerdo Superior N° 007 de 2003 establece los criterios para la clasificación en el escalafón"
                + " del personal docente de la Universidad del Magdalena, aplicables a quienes se encuentren amparados por el Decreto N° 1279 de 2002, o las normas que lo modifiquen o sustituyan."
                + "\n");
                documento.add(p);
                
                
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el artículo 24 del citado Acuerdo, señala los requerimientos para promover a un profesor en la carrera docente."
                + "\n");
        documento.add(p);

        if(!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("AS_ANTIGUO")).equals("N/A")){

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que por su parte, el artículo 27 ibidem señala los requisitos que deben cumplir los profesores nombrados en planta para ser promovidos en el escalafón"
                    + ", indicando para la categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " los siguientes: \n");
            documento.add(p);

            String requisitosxCategoria = "";
            if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asistente")) {
                requisitosxCategoria = "a. Ser profesor Auxiliar de tiempo completo o de medio tiempo y acreditar dos (2) años de permanencia en dicha categoría.\n"
                        + "b. Haber sido evaluado satisfactoriamente en el desempeño de las funciones durante los dos (2) últimos procesos de evaluación docente integral que semestralmente practica la Universidad.\n"
                        + "c. Acreditar mínimo cuarenta (40) horas de formación pedagógica. \n"
                        + "d. Presentar Título de Especialización en el área de titulación Profesional.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asociado")) {
                requisitosxCategoria = "a. Ser profesor Asistente de tiempo completo o de medio tiempo y acreditar "
                        + "dos (2) años de permanencia en dicha categoría.\n"
                        + "b. Haber sido evaluado satisfactoriamente en el desempeño de sus funciones "
                        + "durante los dos (2) últimos procesos de evaluación docente integral que "
                        + "semestralmente practica la Universidad.\n"
                        + "c. Tener título de Maestría.\n"
                        + "d. Someter y obtener la aprobación de pares externos de un (1) trabajo realizado "
                        + "durante su permanencia en la categoría de profesor auxiliar, el cual constituya "
                        + "un aporte significativo a la docencia, a las ciencias, a las artes o a las "
                        + "humanidades de acuerdo con la reglamentación que para el efecto expida el "
                        + "Consejo Superior.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Titular")) {
                requisitosxCategoria = "a. Cumplir los requisitos exigidos para ser profesor asociado.\n"
                        + "b. Acreditar tres (3) años de experiencia calificada en el equivalente de tiempo "
                        + "completo como profesor Asociado.\n"
                        + "c. Haber sido evaluado satisfactoriamente en el desempeño de sus funciones "
                        + "durante los dos últimos procesos de evaluación docente integral que "
                        + "semestralmente practica la Universidad.\n"
                        + "d. Tener título de doctorado.\n"
                        + "e. Someter y obtener la aprobación de pares externos de dos (2) trabajos "
                        + "realizados durante su permanencia en la categoría de profesor asociado que "
                        + "constituyan un aporte significativo a la docencia, a las ciencias, a las artes o "
                        + "a las humanidades de acuerdo con la reglamentación que para el efecto "
                        + "expida el Consejo Superior.\"";
            }

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(fh3c);
            c = new Chunk("\"... ARTÍCULO 27. ",fh3b );
            p.add(c);
            p.add("Los profesores nombrados en planta según la categoría deben cumplir los siguientes requisitos: \n"
            +"... \n");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+"\n",fh3b);
            p.add(c);
            p.add( requisitosxCategoria + "\n");
            documento.add(p);
        }
        else{
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el Artículo 14 del Acuerdo Superior N° 007 de 2021 establece los requisitos mínimos que debe cumplir un profesor de planta"
                + " para ser clasificado al momento de su vinculación en una de las categorías establecidas por el Decreto 1279 de 2002,"
                + " determinando para la categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+" lo siguiente: \n");
                documento.add(p);
            String requisitosxCategoria = "";
            if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asistente")) {
                requisitosxCategoria = "• Título de posgrado.\n" +
                          "• Cinco (5) años de experiencia certificada, en tiempo completo equivalente, en actividades profesorales en programas académicos "
                        + "de educación universitaria o en actividades profesionales afines al perfil en el que va a ser nombrado\n" +
                          "• Productividad académica equivalente a veinte (20) puntos salariales.\"\n";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asociado")) {
                requisitosxCategoria = "• Título de maestría o doctorado.\n" +
                          "• Siete (7) años de experiencia certificada, en tiempo completo equivalente, "
                        + "en actividades profesorales en programas académicos "
                        + "de educación universitaria o en actividades profesionales afines al perfil en el que va a ser nombrado.\n" +
                          "• Productividad académica equivalente a cuarenta (40) puntos salariales.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Titular")) {
                requisitosxCategoria = "• Título de doctorado.\n" +
                          "• Acreditar tres (3) años de experiencia calificada en el equivalente a tiempo completo como profesor asociado"
                        + " de la Universidad del Magdalena.\n" +
                          "• Productividad académica equivalente a ochenta (80) puntos salariales.\"";
            }
            
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(fh3c);
            p.add("\"");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+"\n \n",fh3b);
            p.add(c);
            p.add( requisitosxCategoria + "\n");
            documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que la categoría dentro del escalafón docente es uno de los factores que incide en la modificación de puntos salariales de los docentes de planta,"
                + " de acuerdo con lo establecido en el Literal b. del artículo 12 del Decreto 1279 de 2002.\n"
                + "\n"
                + "Que el artículo 8° de la precitada norma, establece la asignación de puntaje por categoría académica en el escalafón para docentes en carrera, cualquiera que sea su dedicación.\n");
        documento.add(p);

        String categoriaAnterior = "";
        if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asistente")) {
            categoriaAnterior = "Profesor Auxiliar";
        } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asociado")) {
            categoriaAnterior = "Profesor Asistente";
        } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Titular")) {
            categoriaAnterior = "Profesor Asociado";
        }

        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Artículo 17 del referido decreto, establece los criterios a tener en cuenta para la modificación de salario "+
                    "de los docentes que realizan actividades académico-administrativas, disponiéndose, en igual sentido, en el artículo 62 de la mencionada "+
                    "disposición, que el Grupo de Seguimiento al régimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la información a nivel nacional, y además adecuar los criterios "+
                    "y efectuar los ajustes a las metodologías de evaluación aplicadas por los Comités Internos de Asignación de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al Régimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N° 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad académica para los docentes que asuman cargos académico-administrativos, este "+
                    "señaló que independientemente de haber elegido entre la remuneración del cargo que va a desempeñar y la que le corresponde como docente, "+
                    "solo se podrá hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades académico-administrativas, de conformidad con el Artículo 17 del Decreto 1279 de 2002 y el parágrafo 1 del Artículo 60 del Acuerdo N° 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que mediante Resolución Rectoral Nº "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", ", af10b);
            p.add(c);
            p.add("fue comisionado para ejercer un cargo de libre nombramiento y remoción dentro de la Institución como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tomó posesión mediante Acta N° "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }
        
        if(!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("AS_ANTIGUO")).equals("N/A")){
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("solicitó promoción en el escalafón docente de la categoría " + categoriaAnterior + " a la categoría " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " de conformidad con lo prescrito por el Artículo 22 del Acuerdo Superior Nº 007 de 2003.\n");
        documento.add(p);

       } 
        else{
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("solicitó promoción en el escalafón docente de la categoría " + categoriaAnterior + " a la categoría " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " de conformidad con lo prescrito por el Artículo 14 del Acuerdo Superior Nº 007 de 2021.\n");
        documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        try{
        p.add("Que el Consejo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FACULTAD")) + ", en sesión llevada a cabo el día " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0)
                + " contenida en Acta N° " + getNumeroDecimal(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_CONS_FACULTAD"))) + ", estudió la solicitud y emitió concepto favorable de ascenso en el escalafón docente ante Comité Interno de Asignación y Reconocimiento de Puntaje.\n");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        documento.add(p);
        
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        try{
        p.add("Que mediante Acta N° " + listaDatosDocentes.get("ACTA") + " de " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ", tal y como quedo establecido en el ítem " + listaDatosDocentes.get("NUMERAL_ACTA_CIARP")
                +", el Comité Interno de Asignación y Reconocimiento de Puntaje decidió promover en el escalafón docente "+ (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al profesor" : "a la profesora"));
        
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        c = new Chunk(" " +Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("a la categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ", por cumplir con los requisitos fijados para ello y, asignó ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", af10b);
         } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
        p.add(" de acuerdo con el parágrafo II del artículo 8 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que conforme lo dispuesto en el artículo 31 del Acuerdo Superior N° 007 de 2003,"
                + " el Rector expedirá la Resolución de ascenso, previa recomendación motivada del Comité Interno de Asignación y Reconocimiento de Puntaje.\n"
                + "\n"
                + "Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al año el total de puntos que corresponde a cada"
                + " docente, conforme lo señala el Artículo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Promover en el escalafón " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        p.add("a la Categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("Autorizar el reconocimiento y pago de ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", af10b);
        }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get("NOMBRE_SOLICITUD"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                        if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("correspondientes a su promoción de Categoría de " + categoriaAnterior + " a Categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        
        p.add("Los puntos salariales");
                
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p.add(" surtirán efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisión para desempeñar el cargo de libre nombramiento y remoción que le ha sido otorgado.\n");
        }else{
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add(" surtirán efectos fiscales a partir del día "
                + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", fecha en la cual se expidió el acto formal de "+
                "reconocimiento según consta en Acta N° " + listaDatosDocentes.get("ACTA") + ".\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            }
        }
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO QUINTO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEXTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", af10);
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO_JEFE") + "____");
        documento.add(p);
        //</editor-fold>
        return respuesta;
    }

    public Map<String, String> GenerarResolBonificacion(Document documento, String identificacion, String tipo_resolucion) throws BadElementException, IOException, DocumentException, Exception {
        
        int band = 0;
        Double totalPuntos = 0.0;
        Double totalPuntosxActas = 0.0;
        String item = "";
        Map<String, String> datos1 = new HashMap<>();
        List<Map<String, String>> listaDatosDocentes = data_list(3, listaDatos, new String[]{"TIPO_RESOLUCION<->" + tipo_resolucion, "CEDULA<->" + identificacion});
        
        List<Map<String, String>> listaActas = data_list(1, listaDatosDocentes, new String[]{"ACTA"});

        int cantidaddeProductos = listaDatosDocentes.size();
        String[] numeralActas = {};

        //<editor-fold defaultstate="collapsed" desc="estilo">
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
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");

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

        Paragraph p = new Paragraph(10);
        int justificado = Paragraph.ALIGN_JUSTIFIED;
        int centrado = Paragraph.ALIGN_CENTER;
        int derecha = Paragraph.ALIGN_RIGHT;
//</editor-fold>
        
                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">

        //<editor-fold defaultstate="collapsed" desc="HEADER">
        Table headerTable;
        Table headerTableTxt;


        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
         Image imgM = Image.getInstance("C:\\CIARP\\under.png");

        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconoc"
                + "imiento y pago de bonificación por productividad académica"
                + " " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>

        celda = new Cell(new Paragraph(resolucion+ "", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
       
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);
        
        //</editor-fold>

        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(derecha);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
      
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);
        //</editor-fold>

        RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
        
        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="Cuerpo Resolucion">
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;
        
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el Artículo 19 del Decreto N° 1279 de 2002, consagra las bonificaciones como reconocimientos monetarios no salariales, que se reconocen por una sola vez, correspondientes a actividades específicas de productividad académica y no contemplan pagos genéricos indiscriminados.\n"
                + "\n"
                + "Que en el Artículo 20 de la disposición anterior, se establecieron los criterios para el reconocimiento de bonificaciones productividad académica, siendo determinados en igual sentido y en consonancia con esta disposición, en los Artículos 51 y 52 del Acuerdo Superior N° 007 de 2003.\n"
                + "\n"
                + "Que conforme lo establecido en los incisos tercero y cuarto del Artículo 19 del Decreto N° 1279 de 2002, las universidades deben reconocer y pagar semestralmente las bonificaciones que se causen en dichos periodos, de igual forma, las actividades de productividad académica que tengan reconocimiento salarial no reciben bonificaciones.\n");
        documento.add(p);
        
        if(listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE TIEMPO COMPLETO") || 
            listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE MEDIO TIEMPO")    ){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que el Artículo 71 del Acuerdo Superior N° 007 de 19 de marzo de 2003, establece que los Docentes Ocasionales participan de las bonificaciones por productividad académica (Capítulo VI del Decreto 1279 de 19 de junio de 2002), por actividades realizadas durante el periodo en que tienen vinculación.\n");
            documento.add(p);
        }

       
       
        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            totalPuntos += Double.parseDouble(listaDatosDocentes.get(i).get("PUNTOS").replace(",", "."));
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);

        p.add(", presentó solicitud de asignación de puntos de bonificación por ");

        c = new Chunk((listaDatosDocentes.size()==1 ? "un " :numeroEnLetras(listaDatosDocentes.size())) + " (" + listaDatosDocentes.size() + ")", af10b);
        p.add(c);
        p.add((listaDatosDocentes.size()>1 ? " productos, debidamente clasificados y categorizados ": " producto, debidamente clasificado y categorizado "));
        
        
        for (int i = 0; i < listaActas.size(); i++) {
            if (i > 0 && i < listaActas.size() - 1) {
                p.add(", ");
            } else if (i> 0 && i == listaActas.size() - 1) {
                p.add(" y ");
            }
            List<Map<String, String>> listadatosxActasxnumerales = data_list(1, listaDatosDocentes, new String[]{"NUMERAL_ACTA_CIARP"}, new String[]{"ACTA<->" + listaActas.get(i).get("ACTA")});
            for(int j =0 ; j < listadatosxActasxnumerales.size(); j++){
                if (j > 0 && j < listadatosxActasxnumerales.size() - 1) {
                    p.add(", ");
                } else if (j> 0 && j == listadatosxActasxnumerales.size() - 1) {
                    p.add(" y ");
                }
                p.add("en el numeral "+listadatosxActasxnumerales.get(j).get("NUMERAL_ACTA_CIARP"));
                
            }
            try{
            c = new Chunk(" del Acta N° " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", af10);
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            p.add(c);
        }
        

        p.add(".\n");
        documento.add(p);

         

        String norma = "";
        String numerales = "", posicion = "";
        for (int j = 0; j < listaActas.size(); j++) {
            List<Map<String, String>> listadatosxActas = data_list(3, listaDatosDocentes, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA")});
            numerales = "";
            posicion = "";
            List<Map<String, String>> listadatosxActasxTp = data_list(1, listaDatosDocentes, new String[]{"TIPO_PRODUCTO", "NORMA"}, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA")});
            for (int k = 0; k < listadatosxActasxTp.size(); k++) {
                List<Map<String, String>> listadatosxActasTP = data_list(3, listaDatosDocentes, new String[]{"ACTA<->" + listaActas.get(j).get("ACTA"), "NUMERAL_ACTA_CIARP<->" + listadatosxActasxTp.get(k).get("NUMERAL_ACTA_CIARP")});
               
                if (!Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(k).get("NORMA")).equals("#N/D") && !Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(k).get("NORMA")).equals("#N/A")) {
                    try{
                    numerales += (numerales.equals("")?"":", ")+"ítem "+ 
                            
                            getPosicionesNumeral(listadatosxActasTP)+
                           " del numeral "+listadatosxActasxTp.get(k).get("NUMERAL_ACTA_CIARP");
                    }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listadatosxActasxTp.get(k).get("TIPO_PRODUCTO"));
                                        if(Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(j).get("NOMBRE_SOLICITUD")).length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(k).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(k).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                        } 
                
                    norma +=(norma.equals("")?"":", ")+ Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(k).get("NORMA"));
                 
                } else if (Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(k).get("TIPO_PRODUCTO")).equals("Direccion_de_Tesis")) {
                    try{
                    numerales += (numerales.equals("")?"":", ")+"ítem "+ 
                            getPosicionesNumeral(listadatosxActasTP)+
                            " del numeral "+listadatosxActasxTp.get(k).get("NUMERAL_ACTA_CIARP");
                     
                    }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        System.out.println("MENSAJE MENSJAE MEJANE,MASJNUHDJA" + ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listadatosxActasxTp.get(k).get("TIPO_PRODUCTO"));
                                        if(listadatosxActasxTp.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(k).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(k).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                        } 
                    String[] tipoTesis = Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(k).get("NOMBRE_SOLICITUD")).replace("\"", "").split(":");
                    
                    if (tipoTesis[0].equalsIgnoreCase("Tesis de Maestria")) {
                        norma +=(norma.equals("")?"":", ")+ "ítem 1 literal h. numeral II, Artículo 20 del Decreto 1279 del 2002";
                    } else if (tipoTesis[0].equalsIgnoreCase("Tesis de Doctorado")) {
                        norma += (norma.equals("")?"":", ")+ "ítem 2 literal h. numeral II, Artículo 20 del Decreto 1279 del 2002";
                    }
                    
                } else {
                    numerales += (numerales.equals("")?"":", ")+"ítem "+ 
                            getPosicionesNumeral(listadatosxActasTP)+
                            " del numeral "+listadatosxActasxTp.get(k).get("NUMERAL_ACTA_CIARP");
                    norma +=(norma.equals("")?"":", ")+ getNormaProducto(listadatosxActasxTp.get(k).get("TIPO_PRODUCTO"));
                    
                }

            }
            for (int k = 0; k < listadatosxActas.size(); k++) {
                totalPuntosxActas += Double.parseDouble(listadatosxActas.get(k).get("PUNTOS").replace(",", "."));
            }

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            try{
            p.add("Que en virtud de lo anterior y teniendo en cuenta lo señalado en el " + norma + " el Comité Interno de Asignación y Reconocimiento de Puntaje,"
                    + " en sesión realizada el día " + fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0) + " contenida en el Acta N° " + listaActas.get(j).get("ACTA")
                    + " asignó ");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            try{
            c = new Chunk(getNumeroDecimal(""+totalPuntosxActas) + " (" + ValidarNumeroDec(""+totalPuntosxActas) + ") puntos de bonificación ", af10b);
                
            }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listadatosxActasxTp.get(j).get("TIPO_PRODUCTO"));
                                        if(listadatosxActasxTp.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(j).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listadatosxActasxTp.get(j).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
            }
            p.add(c);
            p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al docente " : "a la docente "));
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"))+",", af10b);
            p.add(c);
            p.add(" distribuidos en " + (listaDatosDocentes.size() > 1 ? "los" : "el") + " " + (listaDatosDocentes.size() > 1 ? "productos presentados" : "producto presentado")
                    + ", según el orden establecido en el apartado \"Decisión\" del ");
           try{
            p.add("" +numerales
                +" del Acta N° "+listaActas.get(j).get("ACTA")+" de "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+".\n");
            
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(j).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(j).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }documento.add(p);
            norma="";

            totalPuntosxActas = 0.0;
        }
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al año el total de puntos"
                + " que corresponde a cada docente, conforme lo señala el Artículo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(""+totalPuntos) + " (" + ValidarNumeroDec(""+totalPuntos) + ") puntos de bonificación", af10b);
            
         }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(0).get("TIPO_PRODUCTO"));
                                        if(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
            }
        
        p.add(c);
        p.add(" por la productividad académica que se relaciona en la siguiente tabla:\n");
        documento.add(p);

        ///TABLA
        //<editor-fold defaultstate="collapsed" desc="TABLA">
        float[] tamy = new float[]{5f, 13f, 50f, 8f, 13f, 20f};
        Table TableProductos = new Table(6);
        TableProductos.setWidths(tamy);
        TableProductos.setWidth(100);

        Cell celdaProductos = new Cell(new Paragraph("N°", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Producto", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);
        celdaProductos = new Cell(new Paragraph("Nombre del Producto", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("N° Acta", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Fecha de Acta", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Puntos Reconocidos", af7b));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            celdaProductos = new Cell(new Paragraph("" + (i + 1), af7b));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("TIPO_PRODUCTO")).replace("_", " "), af7b));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            String nombresolicitud = "";
            if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("SUBTIPO_PRODUCTO")).equals("N/A")) {
                nombresolicitud = "" + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("SUBTIPO_PRODUCTO")).replace("_", " ") + ": " + getNOMBREPRODUCTO(listaDatosDocentes.get(i));
            } else {
                nombresolicitud = getNOMBREPRODUCTO(listaDatosDocentes.get(i));
            }
            celdaProductos = new Cell(new Paragraph("" + nombresolicitud, af7));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + listaDatosDocentes.get(i).get("ACTA"), af7));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + fechaEnletras(listaDatosDocentes.get(i).get("FECHA_ACTA"),0), af7));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            
            try{
            celdaProductos = new Cell(new Paragraph("" + getNumeroDecimal(listaDatosDocentes.get(i).get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get(i).get("PUNTOS")) + ") puntos", af7b));
             }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                        if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
            }
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

        }

        documento.add(TableProductos);
        //</editor-fold>

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("\nParágrafo:", af10b);
        p.add(c);
         
        p.add(" Los ");
        try{
        c = new Chunk(getNumeroDecimal(""+totalPuntos) + " (" + ValidarNumeroDec(""+totalPuntos) + ") puntos de bonificación", af10b);
        }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(0).get("TIPO_PRODUCTO"));
                                        if(listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get(0).get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
        }
        p.add(c);
        p.add(" por productividad académica asignados en el recuadro anterior, se establecen según lo dispuesto en el aparte de \"Decisión\" ");

        for (int i = 0; i < listaActas.size(); i++) {
            if (i > 0 && i < listaActas.size() - 1) {
                p.add(", ");
            } else if (i> 0 && i == listaActas.size() - 1) {
                p.add(" y ");
            }
            List<Map<String, String>> listadatosxActasxnumerales = data_list(1, listaDatosDocentes, new String[]{"NUMERAL_ACTA_CIARP"}, new String[]{"ACTA<->" + listaActas.get(i).get("ACTA")});
            for(int j =0 ; j < listadatosxActasxnumerales.size(); j++){
                if (j > 0 && j < listadatosxActasxnumerales.size() - 1) {
                    p.add(", ");
                } else if (j> 0 && j == listadatosxActasxnumerales.size() - 1) {
                    p.add(" y ");
                }
                p.add("en el numeral "+listadatosxActasxnumerales.get(j).get("NUMERAL_ACTA_CIARP"));
                
            }
            try{
            c = new Chunk(" del Acta N° " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", af10);
            
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get(i).get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }p.add(c);
        }
        
        
        p.add(" los cuales, se reconocerán y pagarán por una sola vez" + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? " al" : " a la") + " "+ listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add(", de conformidad con lo dispuesto en el ARTÍCULO 19 del Decreto 1279 de 19 de junio de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", af10);
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        
        p.add(c);
        p.add(listaProyecto.get("REVISO_JEFE") + "____\n");
        documento.add(p);
        //</editor-fold>
        
            return respuesta;
    }

    private Map <String, String> GenerarResolConvalidacion(Document documento, Map<String, String> listaDatosDocentes) throws BadElementException, IOException, DocumentException {
        try{
        
        int band = 0;
        int totalPuntos = 0;
        int totalPuntosxActas = 0;
        String item = "";
        Map<String, String> datos1 = new HashMap<>();
        int cantidaddeProductos = listaDatosDocentes.size();
        String[] numeralActas = {};
        String[] tituloCorto = {};
        tituloCorto = Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")).split(",");

        //<editor-fold defaultstate="collapsed" desc="estilo">
        Font fh1 = new Font(arialFont);
        fh1.setSize(11);
        fh1.setStyle("bold");
        fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("bold");

        Font fh3 = new Font(arialFont);
        fh3.setSize(8);
        fh3.setColor(Color.BLACK);
        
        Font fh3c = new Font(arialFont);
        fh3c.setSize(8);
        fh3c.setStyle("italic");
        fh3c.setColor(Color.BLACK);

        Font fh4 = new Font(arialFont);
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");

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
        int derecha = Paragraph.ALIGN_RIGHT;
//</editor-fold>
        
                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">

        //<editor-fold defaultstate="collapsed" desc="HEADER">

        Table headerTable;
        Table headerTableTxt;


        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
         Image imgM = Image.getInstance("C:\\CIARP\\under.png");

        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconocimiento y pago de puntos salariales por el título de " + tituloCorto[0]
                            + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                            + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        
        celda = new Cell(new Paragraph(resolucion + "", fh4));
        celda.setBorder(0);

        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
       
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);

        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(derecha);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
      
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);
        //</editor-fold>

          RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="Cuerpo Resolucion">
        
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el artículo 62 del Decreto 1279 de junio de 2002, establece que el Grupo de Seguimiento al régimen salarial y prestacional de los profesores universitarios,"
                + " puede definir las directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la información a nivel nacional, y además adecuar los criterios"
                + " y efectuar los ajustes a las metodologías de evaluación aplicadas por los Comités Internos de Asignación de Puntaje o los organismos que hagan sus veces.\n"
                + " \n"
                + "Que el artículo primero, numeral 22 del Acuerdo No. 001 de 4 de marzo de 2004, del Grupo de Seguimiento del régimen salarial y prestacional de los profesores universitarios"
                + " del Decreto 1279 de 19 de junio de 2002, señala sobre la asignación de puntaje por títulos académicos obtenidos en el exterior lo siguiente:\n");
        documento.add(p);
        
        

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setIndentationLeft(13);
        p.setIndentationRight(12);
        p.setFont(fh3c);
        p.add("\"22. El Comité Interno de Asignación y Reconocimiento de Puntaje debe dar trámite a las solicitudes en la medida en que se vayan presentando; las modificaciones salariales tendrán efectos a partir de la fecha en que el Comité expida el acto formal de reconocimiento. \n"
                +"…\n"
                + "En cuanto la asignación de puntaje por títulos académicos obtenidos en el exterior es procedente aplicar lo dispuesto en el Artículo 12 del Decreto 861 de 2000, el cual señala"
                + " “Artículo 12. De los títulos y certificados obtenidos en el exterior requerirán para su validez, de las autenticaciones, registros y equivalencias determinadas por el Ministerio de Educación Nacional"
                + " y el Instituto Colombiano para el Fomento de la Educación Superior.\n"
                + "…\n"
                + "De otro lado si el docente aporta el título para modificación de salario y no cumple con el requisito de la convalidación dentro del plazo señalado, el rector mediante acto administrativo determinará "
                + "que no reconoce los puntos asignados por el incumplimiento de la condición señalada en el acto expedido por el CIARP (Artículo 55 Decreto 1279 de 2002).\n"
                + "\n"
                + "De presentarse la convalidación en el término señalado, en la resolución rectoral que se debe expedir dos veces al año (Artículo 55 Decreto 1279 de 2002) se ordenará el pago del salario modificado, "
                + "desde el momento en que se reconoció y asignó el respectivo puntaje por parte del CIARP, de acuerdo a lo dispuesto en el parágrafo III del Artículo 12 del Decreto 1279,"
                + " el cual dispone: “las modificaciones salariales tienen efecto a partir de la fecha en que el Comité Interno de Asignación y Reconocimiento del Puntaje, "
                + "o el órgano que haga sus veces en cada una de las universidades, expida el acto formal de reconocimiento de los puntos salariales asignados en el marco del presente decreto… ”.\"\n");
        documento.add(p);
        
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Artículo 17 del referido decreto, establece los criterios a tener en cuenta para la modificación de salario "+
                    "de los docentes que realizan actividades académico-administrativas, disponiéndose, en igual sentido, en el Artículo 62 de la mencionada "+
                    "disposición, que el Grupo de Seguimiento al régimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la información a nivel nacional, y además adecuar los criterios "+
                    "y efectuar los ajustes a las metodologías de evaluación aplicadas por los Comités Internos de Asignación de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al Régimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N° 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad académica para los docentes que asuman cargos académico-administrativos, este "+
                    "señaló que independientemente de haber elegido entre la remuneración del cargo que va a desempeñar y la que le corresponde como docente, "+
                    "solo se podrá hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades académico-administrativas, de conformidad con el Artículo 17 del Decreto 1279 de 2002 y el Artículo 6 del Acuerdo N° 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que mediante Resolución Rectoral Nº "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", "+
                    "fue comisionado para ejercer un cargo de libre nombramiento y remoción dentro de la Institución como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tomó posesión mediante Acta N° "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("realizó solicitud de asignación de puntos salariales por el título de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        try{
        p.add("Que el Comité Interno de Asignación y Reconocimiento de Puntaje, en sesión realizada el día " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA_CIARP_INICIAL"),0)
                + " contenida en Acta N° " + listaDatosDocentes.get("ACTA") + ", asignó ");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + getNumero(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", af10b);
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ") + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        try{
        p.add("por el título de " + listaDatosDocentes.get("TITULO") + ", que se pagarán con efectos a partir del "+ fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0)+" al convalidar el título, para lo cual, el docente cuenta hasta el "+ fechaEnletras(listaDatosDocentes.get("FECHA_MAX_CONV"),0)+ ".\n");
       } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        documento.add(p);

        if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RES_TIT_ANT")).equals("N/A")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que mediante resolución rectoral N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RES_TIT_ANT")) + " se reconoce " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ") + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
           try{
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + " " + getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", af10b);
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            p.add(c);
            p.add(" correspondientes al título de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add("Que en dicho acto administrativo se estableció que para la asignación y pago de los puntos, el docente debe cumplir con la convalidación del título dentro de los dos (2) años siguientes, "
                    + "contados a partir del día " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", fecha en la que el Comité Interno de Asignación y Reconocimiento de Puntaje asignó los puntos salariales.\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
                }
                
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el " : "la ") + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add(" presentó ante el Comité Interno de Asignación y Reconocimiento de Puntaje, copia de la " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        try{
        p.add("Que el Comité Interno de Asignación y Reconocimiento de Puntaje, en sesión realizada el día " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                + " contenida en Acta N° " + listaDatosDocentes.get("ACTA") + " punto "+ listaDatosDocentes.get("NUMERAL_ACTA_CIARP")+", determinó tener por cumplido el requisito de la convalidación del título,"
                + " indicando que los ");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        try{
        c = new Chunk(numeroEnLetras(Integer.parseInt(listaDatosDocentes.get("PUNTOS"))) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", af10b);
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
       
        if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
            p.add(" surtirán efectos fiscales desde el " );
            try{
            p.add( fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ".\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        }else{
            p.add(" surtirán efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisión a la que fue "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "encargado" : "encargada")+".\n");
        }
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al año el total de puntos que corresponde a cada"
                + " docente, conforme lo señala el Artículo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", af10b);
        
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }p.add(c);
        p.add(" por el título de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("Los ");
        try{
        c = new Chunk(numeroEnLetras(Integer.parseInt(listaDatosDocentes.get("PUNTOS"))) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", af10b);
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p.add(" surtirán efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisión para desempeñar el cargo de libre nombramiento y remoción que le ha sido otorgado.");
        }else{
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add(" surtirán efectos fiscales a partir del día " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", según consta en Acta N° " + listaDatosDocentes.get("ACTA") + " del " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ".\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
                }
        }
        
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO QUINTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", af10);
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO_JEFE") + "____");
        documento.add(p);
        //</editor-fold>
        
        }catch(Exception e){
            
            e.printStackTrace();
        }
    return respuesta;
    }
    

    private Map <String, String> GenerarResolIngreso(Document documento, Map<String, String> listaDatosDocentes) throws BadElementException, IOException, DocumentException {
        System.out.println("***********ENTRE A INGRESO***************************************");
        int band = 0;
        int totalPuntos = 0;
        int totalPuntosxActas = 0;
        String item = "";
        Map<String, String> datos1 = new HashMap<>();
        int cantidaddeProductos = listaDatosDocentes.size();
        String[] numeralActas = {};

        //<editor-fold defaultstate="collapsed" desc="estilo">
        Font fh1 = new Font(arialFont);
        fh1.setSize(11);
        fh1.setStyle("bold");
        fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("bold");

        Font fh3 = new Font(arialFont);
        fh3.setSize(8);
        fh3.setColor(Color.BLACK);

        Font fh4 = new Font(arialFont);
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");

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
        int derecha = Paragraph.ALIGN_RIGHT;
//</editor-fold>
        
                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">

        //<editor-fold defaultstate="collapsed" desc="HEADER">
RtfHeaderFooterGroup headerDif= new RtfHeaderFooterGroup();
        Table headerTable;
        Table headerTableTxt;


        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
         Image imgM = Image.getInstance("C:\\CIARP\\under.png");

        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el ingreso a la carrera " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        celda = new Cell(new Paragraph(resolucion+"", fh4));
        celda.setBorder(0);

        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
       
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);


        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(derecha);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        //</editor-fold>

          RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="Cuerpo Resolucion">
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que la Ley 30 de 1992 en su Artículo 28 reconoce a las Universidades en virtud del principio de autonomía universitaria,"
                + " entre otros, el derecho a seleccionar a sus profesores.\n");
        documento.add(p);

        if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FA_BECA")).equals("N/A")) {

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que mediante Acuerdo Superior N° 025 de 2002, modificado por el Acuerdo Superior N° 008 de 2014, se adoptó el Programa de Formación Avanzada para la Docencia y la Investigación, "
                    + "disponiendo en su Artículo Séptimo la facultad del Rector de la Universidad del Magdalena para vincular a la planta de personal docente,"
                    + " a los beneficiarios de becas concedidas por organismos o entidades de reconocido prestigio nacional o internacional diferentes a esta institución.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el señor " : "la señora "));
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
            p.add(c);
            p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
            if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
                p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + listaDatosDocentes.get("C_EXPEDICION") + ", ");
            } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
                p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            } else {
                p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            }
            p.add("resultó " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "beneficiario " : "beneficiaria ") + "de un(a) " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FA_BECA"))
                    + ", siendo por ello, vinculado de manera excepcional a la Universidad como Docente de Planta a través del Programa de Formación Avanzada para la Docencia y la Investigación,"
                    + " de conformidad con el Acuerdo Académico N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACUERDO_FA"))+".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que en ese orden, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el señor " : "la señora "));
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
            p.add(c);
            p.add(" fue " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "nombrado " : "nombrada ") + "como " + listaDatosDocentes.get("TIPO_VINCULACION")
                    + " mediante Resolución Rectoral N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_INGRESO")) + ", cargo del cual tomó posesión mediante Acta N° "
                    + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_INGRESO")) + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que luego de haber finalizado los estudios, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " docente en mención fue reincorporado al servicio como " + listaDatosDocentes.get("TIPO_VINCULACION")
                    + " de la Universidad a través de la Resolución Rectoral N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_REINTEGRO")) + ".\n");
            documento.add(p);
        } else if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_INICIO")).equals("N/A")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que mediante Resolución Rectoral No. " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_INICIO")) + ", se dio inicio a la convocatoria pública para proveer cargos docentes en dedicación de Tiempo Completo.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que según Resolución Rectoral N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_FIN")));
            c = new Chunk(" " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
            p.add(c);
            p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
            if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
                p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + listaDatosDocentes.get("C_EXPEDICION") + ", ");
            } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
                p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            } else {
                p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            }
            p.add("resultó " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "ganador" : "ganadora") + " del concurso en el área de desempeño " + listaDatosDocentes.get("AREA_DESEMPEÑO")
                    + " " + listaDatosDocentes.get("FACULTAD") + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que mediante Resolución Rectoral N° " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_INGRESO") )+ " fue " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "nombrado " : "nombrada ")
                    + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el señor " : "la señora "));

            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
            p.add(c);
            p.add(" en el cargo de " + listaDatosDocentes.get("TIPO_VINCULACION") + " de la Universidad del Magdalena, cargo del que tomó posesión mediante Acta N° "
                    + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_INGRESO")) + ".\n");
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el Acuerdo Superior N° 007 de 2003 establece en su Artículo 14, que los profesores aspirantes a la carrera docente se nombrarán de tiempo completo o medio tiempo por un periodo de prueba"
                + " de un (1) año; vencido y aprobado el mismo, el profesor podrá solicitar su ingreso a la carrera ante el Comité de Asignación y Reconocimiento de Puntaje en la categoría que le corresponda en el escalafón.\n"
                + "\n"
                + "Que mediante el Artículo 24 del acuerdo en mención, se establecieron los requisitos para promover a un profesor en la carrera docente, señalándose en igual sentido en el Parágrafo de esta disposición, "
                + "que los profesores con título universitario que cumplan y aprueben el periodo de prueba, ingresan al escalafón docente en la categoría de profesor auxiliar y en la categoría correspondiente para docentes sin título universitario.\n"
                + "\n"
                + "Que por otra parte, el Parágrafo 1° del Artículo 27 del Acuerdo Superior N° 007 de 2003, establece como excepción de la norma antes citada, que el profesor vinculado, escalafonado previamente"
                + " en una universidad pública con un sistema similar al de la Universidad del Magdalena, será ubicado en su categoría, después de superar el periodo de prueba, previa constancia expedida por la universidad de procedencia.\n");
        documento.add(p);

        if (listaDatosDocentes.get("PENDIENTE_INGLES").equalsIgnoreCase("SI")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que mediante Acuerdo Superior N° 008 del 19 de mayo de 2008, se reglamentó el requisito de certificación de la prueba de suficiencia en el manejo del inglés para la provisión de cargos docentes en la Universidad.\n"
                    + "\n"
                    + "Que a través del Acuerdo Superior N° 009 del 03 de mayo de 2013, se modificó el Artículo 1 del Acuerdo Superior N° 008 de 2008, reglamentándose como mínimo un Nivel B2 según el Marco Común de Referencia Europeo"
                    + " para acreditar la suficiencia en el manejo del idioma inglés mediante exámenes con validez internacional o estudios desarrollados en países de habla inglesa.\n"
                    + "\n"
                    + "Que el mencionado acuerdo dispone que el aspirante vinculado a la Universidad cuenta con un máximo de diez (10) meses contados a partir de su vinculación para"
                    + " acreditar el manejo en el idioma inglés so pena que la evaluación de su desempeño en el año de prueba correspondiente sea declarada no satisfactoria.\n"
                    + "\n"
                    + "Que el Acuerdo Superior N° 015 del 30 de noviembre de 2016 modificó el Parágrafo Segundo del artículo Primero del Acuerdo Superior N° 009 de 2013 quedando así:\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(fh3);
            p.add("\"Parágrafo 2: El aspirante que no certifique el mínimo de suficiencia en el manejo del idioma inglés, podrá ser incluido en la lista de elegibles siempre y cuando haya obtenido "
                    + "el puntaje mínimo requerido del total de la calificación final en el proceso de selección. Si el aspirante es vinculado a la universidad, contará con un máximo de hasta veintidós (22)"
                    + " meses contados a partir de su vinculación para acreditar el manejo del idioma inglés so pena de impedirse su ascenso a la categoría de profesor asistente en el escalafón docente\"\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            try{
            p.add("Que el Consejo de " + listaDatosDocentes.get("FACULTAD") + " en sesión celebrada el día " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0) + " contenida en Acta N° "
                    + listaDatosDocentes.get("ACTA_CONS_FACULTAD") + ", determinó que la evaluación del período de prueba " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
            p.add(c);
            p.add("fue superada de manera satisfactoria, realizando la claridad que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " docente debe cumplir con el requisito de manejo del idioma inglés.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
            p.add(c);
            try{
            p.add("solicitó ingreso a la carrera docente ante el Comité Interno de Asignación y Reconocimiento de Puntaje, órgano que en sesión realizada el día " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                    + " contenida en Acta N° " + listaDatosDocentes.get("ACTA") + ", verificó el cumplimiento de los requisitos establecidos para éste y la superación del periodo de prueba, aclarando que tiene pendiente acreditar el manejo del idioma inglés.\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            documento.add(p);
        } else {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            try{
            p.add("Que el Consejo de " + listaDatosDocentes.get("FACULTAD") + " en sesión celebrada el día " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0) + " contenida en Acta N° "
                    + listaDatosDocentes.get("ACTA_CONS_FACULTAD") + " , determinó que la evaluación del período de prueba " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
            p.add(c);
            p.add("fue superada de manera satisfactoria.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
            p.add(c);
            try{
            p.add("solicitó ingreso a la carrera docente ante el Comité Interno de Asignación y Reconocimiento de Puntaje, órgano que en sesión realizada el día " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                    + " contenida en Acta N° " + listaDatosDocentes.get("ACTA") + ", verificó el cumplimiento de los requisitos establecidos para éste y la superación del periodo de prueba.\n");
            } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Autorizar el ingreso a la carrera " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        p.add("en la categoría de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CAT_INGRESO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("El ingreso a la carrera "+ (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
            try{
        p.add(", producirán efectos fiscales a partir del " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", fecha en la que el Comité verificó el cumplimiento de los requisitos, según consta en Acta N° " + listaDatosDocentes.get("ACTA") + " del " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ".\n");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        }
        documento.add(p);
       
        if(listaDatosDocentes.get("PENDIENTE_INGLES").equalsIgnoreCase("SI")){
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        p.add("Para acreditar el manejo del idioma inglés, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("tiene un plazo de veintidós (22) meses contados a partir de la fecha de reincorporación que se produjo mediante Resolución Rectoral N° "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_REINTEGRO"))
        +", so pena de impedirse su ascenso a la categoría de profesor asistente, conforme lo señala el artículo 1 del Acuerdo Superior N° 015 de 2016.\n");
        documento.add(p);
           
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO QUINTO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEXTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);
       }else{
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", af10b);
        p.add(c);
        p.add("Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + "el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO QUINTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);
        }
       
        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", af10);
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO_JEFE") + "____");
        documento.add(p);
        //</editor-fold>
       return respuesta; 
    }

    private Map<String, String> GenerarResolTitulacion(Document documento, Map<String, String> listaDatosDocentes) throws BadElementException, IOException, DocumentException, Exception {
        System.out.println("***********ENTRE A TITULACION");
        int band = 0;
        int totalPuntos = 0;
        int totalPuntosxActas = 0;
        String item = "";
        Map<String, String> datos1 = new HashMap<>();
        int cantidaddeProductos = listaDatosDocentes.size();
        String[] numeralActas = {};
        String[] tituloCorto = {};
        tituloCorto = Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).split(",");

        //<editor-fold defaultstate="collapsed" desc="estilo">
        Font fh1 = new Font(arialFont);
        fh1.setSize(11);
        fh1.setStyle("bold");
        fh1.setStyle("underlined");

        Font fh2 = new Font(arialFont);
        fh2.setSize(11);
        fh2.setColor(Color.BLACK);
        fh2.setStyle("bold");

        Font fh3 = new Font(arialFont);
        fh3.setSize(8);
        fh3.setColor(Color.BLACK);
        
        Font fh4 = new Font(arialFont);
        fh4.setSize(8);
        fh4.setColor(Color.BLACK);
        fh4.setStyle("bold");

        Font fh3c = new Font(arialFont);
        fh3c.setSize(9);
        fh3c.setColor(Color.BLACK);
        fh3c.setStyle("italic");
             
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
        int derecha = Paragraph.ALIGN_RIGHT;
//</editor-fold>
        System.out.println("''''''''''''entre a encabezado");
                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ENCABEZADO Y EL PIE DE PAGINA">

        //<editor-fold defaultstate="collapsed" desc="HEADER">
        RtfHeaderFooterGroup headerDif= new RtfHeaderFooterGroup();
       
       
        Table headerTable;
        Table headerTableTxt;
       


        Image imgL = Image.getInstance("C:\\CIARP\\escudo.png");
        Image imgM = Image.getInstance("C:\\CIARP\\under.png");
       
        headerTable = new Table(1, 2);
        headerTable.setWidth(100);

        Cell celda = new Cell(imgL);
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCIÓN N°\n", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconocimiento y pago de puntos salariales por el título de " + tituloCorto[0]
                + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        
        celda = new Cell(new Paragraph(resolucion+"", fh4));
        celda.setBorder(0);

        celda.setHorizontalAlignment(Cell.ALIGN_JUSTIFIED);
        headerTable.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTable.addCell(celda);
       
        headerTableTxt = new Table(1, 1);
        headerTableTxt.setWidth(100);
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTORÍA – Resolución Nº ", fh4));
        celda.setBorder(0);
        headerTableTxt.addCell(celda);
        celda = new Cell(imgM);
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_TOP);
        celda.setHorizontalAlignment(Cell.ALIGN_LEFT);
        headerTableTxt.addCell(celda);
       

        //</editor-fold>
       
        //<editor-fold defaultstate="collapsed" desc="FOOTER">
        float[] tam = new float[4];
        tam[0] = 7;//texto
        tam[1] = 1;// num page
        tam[2] = 1.5f;//slide
        tam[3] = 1;// num pages

        Table footertable = new Table(4, 1);
        footertable.setWidth(20);
        footertable.setAlignment(derecha);

        footertable.setWidths(tam);

        celda = new Cell(new Paragraph("Página ", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(fh4));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(fh4));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        //</editor-fold>
       
        RtfHeaderFooterGroup headerg = new RtfHeaderFooterGroup();
        RtfHeaderFooter headeresc = new RtfHeaderFooter(headerTable);
        RtfHeaderFooter headertxt = new RtfHeaderFooter(headerTableTxt);
       
        headerg.setHeaderFooter(headeresc, RtfHeaderFooter.DISPLAY_FIRST_PAGE);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_LEFT_PAGES);
        headerg.setHeaderFooter(headertxt, RtfHeaderFooter.DISPLAY_RIGHT_PAGES);

       
        RtfHeaderFooter footer = new RtfHeaderFooter(footertable);

        documento.setHeader(headerg);
        documento.setFooter(footer);
        //</editor-fold>
        //<editor-fold defaultstate="collapsed" desc="Cuerpo Resolucion">
        p = new Paragraph(10);
        justificado = Paragraph.ALIGN_JUSTIFIED;
        centrado = Paragraph.ALIGN_CENTER;

        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("“UNIMAGDALENA”", af10b);
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el Capítulo III del Decreto 1279 de 19 de junio de 2002, establece los factores y criterios que inciden en la modificación de puntos salariales "
                + "de los docentes amparados por este régimen.\n"
                + "\n"
                + "Que los títulos correspondientes a estudios universitarios de pregrado o posgrado es uno de los factores que incide en las modificaciones de los puntos"
                + " salariales de los docentes de planta, de acuerdo con lo establecido en el Literal a. del artículo 12 del Decreto 1279 del 19 de junio de 2002.\n");
        documento.add(p);

        String titulo = "";
        String literal ="";
        
        if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Doctorado")) {
            titulo = "PhD. o Doctorado";
            literal ="artículo 7 Numeral 2, Literal c del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Maestría")) {
            titulo = "Magister o Maestría";
            literal ="artículo 7 Numeral 2, Literal b del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Especialización")) {
            titulo = "Especialización";
            literal ="artículo 7 Numeral 2, Literal a del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Especialización Clínica")) {
            titulo = "Especialización Clínica";
            literal ="artículo 7 Numeral 2, parágrafo II del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Pregrado")) {
            titulo = "Pregrado";
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que el "+literal+" establece la asignación de puntaje por títulos universitarios de posgrado,"
                + " disponiendo para el título de " + titulo + " lo siguiente: \n");
        documento.add(p);

        String requisitosxTitulo = "";
        if (titulo.equals("PhD. o Doctorado")) {
            requisitosxTitulo = "c. Por título de Ph. D. o Doctorado equivalente se asignan hasta ochenta (80) puntos. Cuando el docente "
                    + "acredite un título de Doctorado, y no tenga ningún título acreditado de Maestría, se le otorgan hasta "
                    + "ciento veinte (120) puntos. No se conceden puntos por títulos de Magister o Maestría posteriores al "
                    + "reconocimiento de ese doctorado.";
        } else if (titulo.equals("Magister o Maestría")) {
            requisitosxTitulo = "b. Por el título de Magister o Maestría se asignan hasta cuarenta (40) puntos.";
        } else if (titulo.equals("Especialización")) {
            requisitosxTitulo = "a. Por títulos de Especialización cuya duración esté entre uno (1) y dos (2) años académicos, hasta veinte "
                    + "(20) puntos. Por año adicional se adjudican hasta diez (10) puntos hasta completar un máximo de "
                    + "treinta (30) puntos. Cuando el docente acredite dos (2) especializaciones se computa el número de "
                    + "años académicos y se aplica lo señalado en este literal. No se reconocen más de dos (2) "
                    + "especializaciones.";
        } else if (titulo.equals("Especialización Clínica")) {
            requisitosxTitulo = "PARÁGRAFO II. Para el caso de las especializaciones clínicas en medicina humana y odontología, se "
                    + "adjudican quince (15) puntos por cada año, hasta un máximo acumulable de setenta y cinco (75) puntos.";
        } else if (titulo.equals("Pregrado")) {
            requisitosxTitulo = "a. Por título de pregrado, ciento setenta y ocho (178) puntos.\n"
                    + "b. Por título de pregrado en medicina humana o composición musical, ciento ochenta y tres (183) puntos.";
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setIndentationLeft(13);
        p.setIndentationRight(12);
        p.setFont(fh3c);
        p.add("\"" + requisitosxTitulo + "\"\n");
        documento.add(p);
        
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Artículo 17 del referido decreto, establece los criterios a tener en cuenta para la modificación de salario "+
                    "de los docentes que realizan actividades académico-administrativas, disponiéndose, en igual sentido, en el Artículo 62 de la mencionada "+
                    "disposición, que el Grupo de Seguimiento al régimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la información a nivel nacional, y además adecuar los criterios "+
                    "y efectuar los ajustes a las metodologías de evaluación aplicadas por los Comités Internos de Asignación de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al Régimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N° 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad académica para los docentes que asuman cargos académico-administrativos, este "+
                    "señaló que independientemente de haber elegido entre la remuneración del cargo que va a desempeñar y la que le corresponde como docente, "+
                    "solo se podrá hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades académico-administrativas, de conformidad con el Artículo 17 del Decreto 1279 de 2002 y el parágrafo 1 del Artículo 60 del Acuerdo N° 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(af10);
            p.setAlignment(justificado);
            p.add("Que mediante Resolución Rectoral Nº "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", "+
                    "fue comisionado para ejercer un cargo de libre nombramiento y remoción dentro de la Institución como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tomó posesión mediante Acta N° "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }

        if (titulo.equals("PhD. o Doctorado") && !Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO_MAEST_PUNTAJE")).equals("N/A")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(af10);
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " docente ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
            p.add(c);
            p.add("se le asignaron cuarenta (40) puntos salariales por el título de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO_MAEST_PUNTAJE"))+ ".\n");
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("realizó solicitud de asignación de puntos salariales por titulación, aportando copia del diploma de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        try{
        p.add("Que el Comité Interno de Asignación y Reconocimiento de Puntaje, en sesión realizada el día " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                + " contenida en el numeral "+ listaDatosDocentes.get("NUMERAL_ACTA_CIARP") +" del Acta N° " + listaDatosDocentes.get("ACTA") + ", asignó ");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", af10b);
        }catch (Exception ex) {
                                    Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                        respuesta.put("ESTADO", "ERROR");
                                        respuesta.put("MENSAJE", ""+ex.getMessage());
                                        respuesta.put("LINEA_ERROR_DOCENTE", ""+listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"));
                                        respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                        if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get("NOMBRE_SOLICITUD"));
                                        }else{
                                        respuesta.put("NOMBRE_PRODUCTO", ""+listaDatosDocentes.get("NOMBRE_SOLICITUD").substring(0,100));
                                        }
                                        return respuesta;  
                                }
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ") + listaDatosDocentes.get("TIPO_VINCULACION")+" ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", af10b);
        p.add(c);
        p.add("por el título de " + Utilidades.Utilidades.decodificarElemento(tituloCorto[0]) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al año el total de puntos que corresponde a cada"
                + " docente, conforme lo señala el Artículo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add("En mérito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO PRIMERO: ", af10b);
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), af10b);
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("cédula de ciudadanía N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("cédula de extranjería N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N° " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", af10b);
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);  
        p.add(" por el título de " + Utilidades.Utilidades.decodificarElemento(tituloCorto[0]).replace("\"", "") + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO SEGUNDO: ", af10b);
        p.add(c);
        p.add("Los ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", af10b);
       } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
        p.add(c);
        
       
        
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p.add(" surtirán efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisión para desempeñar el cargo de libre nombramiento y remoción que le ha sido otorgado.\n");
        }else{
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add(" surtirán efectos fiscales a partir del día " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", según consta en Acta N° " + listaDatosDocentes.get("ACTA") + ".\n");
        } catch (Exception ex) {
                                Logger.getLogger(GeneracionCartas.class.getName()).log(Level.SEVERE, null, ex);
                                respuesta.put("ESTADO", "ERROR");
                                respuesta.put("MENSAJE", ""+ex.getMessage());
                                respuesta.put("LINEA_ERROR_DOCENTE", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")));
                                respuesta.put("LINEA_ERROR_PRODUCTO", ""+listaDatosDocentes.get("TIPO_PRODUCTO"));
                                 if(listaDatosDocentes.get("NOMBRE_SOLICITUD").length()<=100){
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")));
                                    }else{
                                    respuesta.put("NOMBRE_PRODUCTO", ""+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).substring(0,100));
                                    }
                                return respuesta;
                            }
                }
        }
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO TERCERO: ", af10b);
        p.add(c);
        p.add("Notificar el contenido de la presente decisión " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".", af10b);
        p.add(c);
        p.add(" Para tal efecto envíesele la correspondiente comunicación al correo electrónico de notificación, haciéndole saber que contra la misma procede"
                + " recurso de reposición ante el mismo funcionario que la expide, el cual deberá presentar y sustentar por escrito dentro de los diez (10) días"
                + " siguientes a la notificación, de acuerdo con lo preceptuado en los Artículos 76 y 77 de la Ley 1437 de 2011, "
                + " el Artículo 55 del Decreto 1279 de 2002 y el Artículo 47 del Acuerdo Superior N° 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO CUARTO:", af10b);
        p.add(c);
        p.add(" Comunicar la presente decisión a la Dirección de Talento Humano y al Grupo de Nómina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el trámite correspondiente. Envíese copia de esta resolución al Comité Interno de Asignación y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", af10b);
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        c = new Chunk("ARTÍCULO QUINTO:", af10b);
        p.add(c);
        p.add(" La presente resolución rige a partir del término de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add("NOTIFÍQUESE, COMUNÍQUESE Y CÚMPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(af10);
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(af10b);
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", af10);
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Proyectó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(af7);
        c = new Chunk("Revisó: ", af7b);
        p.add(c);
        p.add(listaProyecto.get("REVISO_JEFE") + "____");
        documento.add(p);
        //</editor-fold>
        return respuesta;
    }

    private String numeroEnLetras(int numero) throws Exception{
        String[] Unidades, Decenas, Centenas;
        String Resultado = "";

        /**
         * ************************************************
         * Nombre de los números
         * ************************************************
         */
        Unidades = new String[]{"", "Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veintiún", "Veintidós", "Veintitrés", "Veinticuatro", "Veinticinco", "Veintiséis", "Veintisiete", "Veintiocho", "Veintinueve"};
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

    private String getNormaProducto(String tipoProducto) {
        String retNorma = "";
        for (int i = 0; i < normaProducto.size(); i++) {
            if (normaProducto.get(i).get("PRODUCTO").equals(tipoProducto)) {
                retNorma = normaProducto.get(i).get("NORMA");
                break;
            }
        }
        return retNorma;
    }

    public String getNorma(Map<String, String> datos) {
        String ret = "";

        if (datos.get("").equals("")) {

        }

        return ret;
    }

    private String getNumeroDecimal(String numero) throws Exception{
        String retorno = "";
        
        if(numero.indexOf(",") > -1){
            
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

    private String ValidarNumero(String numero) throws Exception {
        return (numero.equals("N/A") ? "0" : numero);
    }

    private String fechaEnletras(String fecha, int opc) throws Exception{// 7/08/2012

        String fechaletra = "";
        if (!fecha.equals("N/A")) {
            String[] dividirFecha = fecha.split("/");
            String[] meses = {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"};

            String dia = numeroEnLetras(Integer.parseInt(dividirFecha[0]));
            String mes = meses[Integer.parseInt(dividirFecha[1]) - 1];

            fechaletra = dividirFecha[0] + " de " + mes + " de " + dividirFecha[2];
            if (opc == 1) {
                fechaletra = dia + " (" + dividirFecha[0] + ") dias del mes de " + mes + " de " + dividirFecha[2];
            }
        }

        return fechaletra;
    }

    private String getNOMBREPRODUCTO(Map<String, String> datos) {
        String datosProducto = "";
        switch (datos.get("TIPO_PRODUCTO")) {

            case "Articulo":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la revista " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

                case "Art_Col":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la revista " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Libro":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD").replace("\"", ""))
                            + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Capitulo_de_Libro":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; editorial " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISBN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Ponencias_en_Eventos_Especializados":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; en el " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Publicaciones_Impresas_Universitarias":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN/ISBN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Reseñas_Críticas":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la revista " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Traducciones":
            {
                try {
                    datosProducto = " \"" + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + "\"; de la revista " + Utilidades.Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                            + "; ISSN: " + Utilidades.Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                            + "; " + Utilidades.Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                            + " (" + Utilidades.Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                            + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
                break;

            case "Direccion_de_Tesis":
                datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "") + " ";
                break;
            default:
                if (!Utilidades.Utilidades.decodificarElemento(datos.get("N_AUTORES")).equals("N/A")) {
            try {
                datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                        + ", " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
            } catch (Exception ex) {
                Logger.getLogger(GeneracionResoluciones.class.getName()).log(Level.SEVERE, null, ex);
            }
                } else {
                    datosProducto = " " + Utilidades.Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")).replace("\"", "")
                            + " ";
                }
                break;
        }

        return datosProducto;
    }

    private List<Map<String, String>> getTipoProductoSalarial(List<Map<String, String>> lista) {
        List<Map<String, String>> retorno = new ArrayList<>();
        int ban = 0;
        for (int i = 0; i < lista.size(); i++) {
            ban = 0;
            if (lista.get(i).get("TIPO_PRODUCTO").equals("Convalidacion")) {
                ban = 1;
            } else if (lista.get(i).get("TIPO_PRODUCTO").equals("Titulacion")) {
                ban = 1;
            } else if (lista.get(i).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")) {
                ban = 1;
            }

            if (ban == 0) {
                retorno.add(lista.get(i));
            }
        }

        return retorno;
    }

    private Double getSumaPuntos(List<Map<String, String>> listadatosxActas) {
        Double suma=0.0;
        DecimalFormat formateador = new DecimalFormat("#.#");
        
        for(Map<String, String> datos:listadatosxActas){
            
            suma +=Double.parseDouble(datos.get("PUNTOS").replace(",", "."));
        }
        
        suma = Double.parseDouble(formateador.format(suma).replace(",", "."));
        
        return suma;
    }
    
    private String numeroOrdinales(int numero) throws Exception{
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

        }
        return Resultado.toLowerCase();
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

    private String getNombreNumero(int numero, String articulo) throws Exception {
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

    private String getNumero(String numero) {
        String retorno = "";
        if(numero.indexOf(",") > -1){
            
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
                retorno = numrs[0];
                if(Integer.parseInt(numrs[1]) > 0){
                    retorno += ","+ numrs[1];//
                }
            } else {
                retorno = numero;
            }
        }
        return retorno;
    }

    private String getPosicionesNumeral(List<Map<String, String>> listadatosxActasTP) throws Exception {
        String ret = "";
        for (int i = 0; i < listadatosxActasTP.size(); i++) {
            if (i > 0 && i < listadatosxActasTP.size() - 1) {
                ret += ", ";
            } else if (i> 0 && i == listadatosxActasTP.size() - 1) {
                ret += " y ";
            }
            
            ret += numeroOrdinales(Integer.parseInt(ValidarNumeroDec(Utilidades.Utilidades.decodificarElemento(listadatosxActasTP.get(i).get("POSICION_ACTA")).equals("N/A")?"0":""+listadatosxActasTP.get(i).get("POSICION_ACTA"))));
            
        } 
        
        return ret;
        
    }

    private List<Map<String, String>> getSinTipoResolucion(List<Map<String, String>> listaTipoResolucion, List<Map<String, String>> listaCargoAcademicoAdministrativoxtipoProducto) {
        List<Map<String, String>> listaretorno = new ArrayList<>();
        boolean encontro = false;
        for(Map<String, String> carg_acad: listaCargoAcademicoAdministrativoxtipoProducto){
            String tiporelso = getResolucionxProducto(carg_acad.get("TIPO_PRODUCTO"));
            encontro=false;
            for(Map<String, String> tiposRel:listaTipoResolucion){
                if(tiposRel.get("TIPO_RESOLUCION").equals(tiporelso)){
                    encontro = true;
                    break;
                }
            }
            if(!encontro){
                listaretorno.add(carg_acad);
            }
        }
        
        return listaretorno;
    }

    private String getResolucionxProducto(String tipoProducto) {
        String resolucion = "";
        if (tipoProducto.toUpperCase().equals("Convalidacion".toUpperCase())) {
            resolucion = "Convalidacion";
        } else if (tipoProducto.toUpperCase().equals("Titulacion".toUpperCase())) {
            resolucion = "Titulacion";
        } else if (tipoProducto.toUpperCase().equals("Ascenso_en_el_Escalafon_Docente".toUpperCase())) {
            resolucion = "Ascenso_en_el_escalafon";
        } else{
            resolucion = "Salarial";
        }   
        return resolucion;
    }
   
    public String ValidarNumeroDec(String valor)throws Exception{

 

        String retorno = "";
        
        if(valor.indexOf(",") > -1){
            
            valor = valor.replace(",", ".");
        }
        else{
            retorno =valor;
        }
        
        if (valor.indexOf(".") > -1) {
            
            
            Double dat = Double.parseDouble(valor);
            
            DecimalFormat df = new DecimalFormat("0.0");
            
            valor = df.format(dat);
            
            valor = valor.replace(".", ",");
            String[] daot= valor.split(",");
            

            if(daot[1].equals("0")){
                retorno = daot[0];
            }else{
                retorno = valor;
            }
            
        }
        return retorno;
 

    }
     
}
