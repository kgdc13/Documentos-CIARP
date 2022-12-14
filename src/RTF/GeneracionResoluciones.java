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
        auxiliarJerarquia.put("NORMA", "Art??culo 14 del Acuerdo Superior N?? 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ascenso_en_el_Escalafon_Docente");
        auxiliarJerarquia.put("NPRODUCTO", "Ascenso en el Escalaf??n Docente");
        auxiliarJerarquia.put("NORMA", "Art??culo 27 del Acuerdo Superior N?? 007 de 2003");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Titulacion");
        auxiliarJerarquia.put("NPRODUCTO", "Titulaci??n");
        auxiliarJerarquia.put("NORMA", "Art??culo 7 del Decreto 1279 del 2002, Art??culo Primero del Acuerdo 001 de 2004 del Grupo de Seguimiento al Decreto 1279 de 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Articulo");
        auxiliarJerarquia.put("NPRODUCTO", "Art??culo");
        auxiliarJerarquia.put("NORMA", "literal a. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Video_Cinematograficas_o_Fonograficas");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n de Video Cinematografica o Fonografica");
        auxiliarJerarquia.put("NORMA", "literal b. numeral I, Art??culo 10; literal a. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Libro");
        auxiliarJerarquia.put("NORMA", "literales c, d, e, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Capitulo_de_Libro");
        auxiliarJerarquia.put("NPRODUCTO", "Cap??tulo de Libro");
        auxiliarJerarquia.put("NORMA", "literales c, d, e, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Premios_Nacionales_e_Internacionales");
        auxiliarJerarquia.put("NPRODUCTO", "Premio Nacional o Internacional");
        auxiliarJerarquia.put("NORMA", "literal f. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Patente");
        auxiliarJerarquia.put("NPRODUCTO", "Patente");
        auxiliarJerarquia.put("NORMA", "literal g. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traduccion_de_Libros");
        auxiliarJerarquia.put("NPRODUCTO", "Traducci??n de Libro");
        auxiliarJerarquia.put("NORMA", "literal h. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Obra_Artistica");
        auxiliarJerarquia.put("NPRODUCTO", "Obra Artistica");
        auxiliarJerarquia.put("NORMA", "literal i. numeral I, Art??culo 10, literal g. numeral II, Articulo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_Tecnica");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n T??cnica");
        auxiliarJerarquia.put("NORMA", "literal j. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Produccion_de_Software");
        auxiliarJerarquia.put("NPRODUCTO", "Producci??n de Software");
        auxiliarJerarquia.put("NORMA", "literal k. numeral I, Art??culo 10 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Ponencias_en_Eventos_Especializados");
        auxiliarJerarquia.put("NPRODUCTO", "Ponencia en Evento Especializado");
        auxiliarJerarquia.put("NORMA", "literal b. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Publicaciones_Impresas_Universitarias");
        auxiliarJerarquia.put("NPRODUCTO", "Publicaci??n Impresa Universitaria");
        auxiliarJerarquia.put("NORMA", "literal c. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Estudios_Posdoctorales");
        auxiliarJerarquia.put("NPRODUCTO", "Estudio Posdoctoral");
        auxiliarJerarquia.put("NORMA", "literal d. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "el");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Rese??as_Cr??ticas");
        auxiliarJerarquia.put("NPRODUCTO", "Rese??a Cr??tica");
        auxiliarJerarquia.put("NORMA", "literal e. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Traducciones");
        auxiliarJerarquia.put("NPRODUCTO", "Traducci??n");
        auxiliarJerarquia.put("NORMA", ""
                + "literal f. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Direccion_de_Tesis");
        auxiliarJerarquia.put("NPRODUCTO", "Direcci??n de Tesis");
        auxiliarJerarquia.put("NORMA", "literal h. numeral II, Art??culo 20 del Decreto 1279 del 2002");
        auxiliarJerarquia.put("ARTICULO", "la");
        normaProducto.add(auxiliarJerarquia);

        auxiliarJerarquia = new HashMap<>();
        auxiliarJerarquia.put("PRODUCTO", "Evaluacion_como_par");
        auxiliarJerarquia.put("NPRODUCTO", "Evaluaci??n como par");
        auxiliarJerarquia.put("NORMA", "literal i. numeral I, Art??culo 20 del Decreto 1279 del 2002");
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
       
        Fonts f = new Fonts(arialFont);
        
        

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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "???Por la cual se autoriza el reconocimiento y pago de puntos salariales "
                + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " "
                + NOM_DOCENTE + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>

        celda = new Cell(new Paragraph(resolucion+ "???", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
     
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.setAlignment(centrado);
        c = new Chunk("CONSIDERANDO:\n", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        if(listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE TIEMPO COMPLETO") || 
            listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE MEDIO TIEMPO")    ){
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que el Art??culo 70 del Acuerdo Superior N?? 007 de 2003 establece las condiciones para la valoraci??n de la productividad acad??mica de los docentes ocasionales. \n");
            documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.setAlignment(justificado);
        p.add("Que el Cap??tulo III del Decreto N?? 1279 de 2002, establece los factores y criterios a tener en cuenta para la modificaci??n de puntos salariales de los docentes amparados por dicho r??gimen, siendo la productividad acad??mica, uno de los factores incidentes en este proceso, seg??n lo establecido en el Literal C., del Art??culo 12 y el Art??culo 16 de la disposici??n en cita.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.setAlignment(justificado);
        p.add("Que el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje ??? CIARP, ha considerado para la asignaci??n de puntaje de los docentes de planta, los criterios de evaluaci??n y asignaci??n que para el efecto ha establecido el Grupo de Seguimiento del R??gimen Salarial y Prestacional de los Profesores Universitarios mediante el Acuerdo N?? 001 de 04 de marzo de 2004.\n");
        documento.add(p);
        
        if(listaDatosDocentesCargoAcad.size()>0){
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Art??culo 17 del referido decreto, establece los criterios a tener en cuenta para la modificaci??n de salario "+
                    "de los docentes que realizan actividades acad??mico-administrativas, disponi??ndose, en igual sentido, en el Art??culo 62 de la mencionada "+
                    "disposici??n, que el Grupo de Seguimiento al r??gimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la informaci??n a nivel nacional, y adem??s adecuar los criterios "+
                    "y efectuar los ajustes a las metodolog??as de evaluaci??n aplicadas por los Comit??s Internos de Asignaci??n de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al R??gimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N?? 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad acad??mica para los docentes que asuman cargos acad??mico-administrativos, este "+
                    "se??al?? que independientemente de haber elegido entre la remuneraci??n del cargo que va a desempe??ar y la que le corresponde como docente, "+
                    "solo se podr?? hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades acad??mico-administrativas, de conformidad con el Art??culo 17 del Decreto 1279 de 2002 y el par??grafo 1 del Art??culo 60 del Acuerdo N?? 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que mediante Resoluci??n Rectoral N?? "+
                    listaDatosDocentes.get(0).get("RESOL_ENCARGO")+" "+
                    (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get(0).get("TIPO_VINCULACION"));
            c = new Chunk(" "+Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE)+", ",Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
                    
            p.add("fue "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "comisionado" : "comisionada")+" para ejercer un cargo de libre nombramiento y remoci??n dentro de la Instituci??n como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tom?? posesi??n mediante Acta N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", present?? solicitud de asignaci??n de puntos salariales por" + (listaDatosDocentes.size()==1?" un":numeroEnLetras(listaDatosDocentes.size())) + " (" + listaDatosDocentes.size() + ")");
        
        
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
                p.add("en el ??tem "+listadatosxActasxnumerales.get(j).get("NUMERAL_ACTA_CIARP"));
            }
            try{
            p.add(" del Acta N?? " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "");
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
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            if(listaDatosDocentesAux.get(0).get("RETROACTIVIDAD").length()==10){
                try{
            p.add("Que la base de datos de COLCIENCIAS para Revistas Internacionales fue actualizada el d??a " + fechaEnletras(listaDatosDocentesAux.get(0).get("RETROACTIVIDAD"),0) + ", seg??n la p??gina web www.colciencias.gov.co.\n");
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
                    numerales += (numerales.equals("")?"":", ")+"??tem "+ 
                            getPosicionesNumeral(listadatosxActasTP)+
                      
                            " del numeral "+listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP");
                    norma += (norma.equals("")?"":", ")+Utilidades.Utilidades.decodificarElemento(listadatosxActasxTp.get(h).get("NORMA"));
                } else {
                    if (!listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ARTICULO")
                            && !listadatosxActasxTp.get(h).get("TIPO_PRODUCTO").toUpperCase().equals("ART_COL")) {
                        numerales += (numerales.equals("")?"":", ")+"??tem "+ 
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
                                norm += "??tem B, literal a. A.1";
                                }else{
                                        norm += "??tem A, literal a. A.1";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("A2")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "??tem B, literal a. A.2";
                                }else{
                                        norm += "??tem A, literal a. A.2";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("B")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "??tem B, literal a. A.3";
                                }else{
                                        norm += "??tem A, literal a. A.3";
                                        }
                            } else if (Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("CATEGORIA")).equals("C")) {
                                if( Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Revision de tema")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Articulo corto")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Reportes de caso")||
                                        Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Carta al editor")||
                                       Utilidades.Utilidades.decodificarElemento(listaxcategoria.get(hh).get("SUBTIPO_PRODUCTO")).equals("Editorial")){
                                norm += "??tem B, literal a. A.4";
                                }else{
                                        norm += "??tem A, literal a. A.4";
                                        }
                            }
                        }
                        num+= "??tem "+ getPosicionesNumeral(listadatosxActasTP);
                        num += " del numeral "+listadatosxActasxTp.get(h).get("NUMERAL_ACTA_CIARP");
                        
                        norm += ", art??culo 10 del Decreto 1279";
                        norma += (norma.equals("")?"":", ")+norm;
                        numerales += (numerales.equals("")?"":", ")+num;    
                    }
                }
            }
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);

            String numeroletras = getNumeroDecimal(""+listadatosxActas.size());
            
            try{
            p.add("Que en virtud de lo anterior y conforme a lo se??alado en "
                + norma
            
                +" el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje en sesi??n realizada el d??a "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+", "
                +"mediante Acta N?? "+listaActas.get(j).get("ACTA")+" decidi?? asignar "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al docente " : "a la docente "));
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
            c = new Chunk(NOM_DOCENTE+ " ",Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            try{
            p.add(
                getNumeroDecimal(""+totalp)+" ("+ValidarNumeroDec(""+totalp)+")"
               
                +" puntos salariales, "
                +"distribuidos en " +(listadatosxActas.size()>1 ? "los " + numeroletras:((listadatosxActas.size()== 1?numeroletras.substring(0, 2):numeroletras)))
                    
                    +" ("+listadatosxActas.size()+") "+(listadatosxActas.size()>1 ? "productos":"producto")+" seg??n el orden establecido en el aparte de ???Decisi??n??? del "
                +numerales +" del Acta N?? "+listaActas.get(j).get("ACTA")+" de "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+".\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.setAlignment(justificado);
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA, se determina dos (2) veces al a??o el total de puntos que corresponde a cada docente, conforme lo se??ala el Art??culo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);
        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.setAlignment(justificado);
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.setAlignment(centrado);
        c = new Chunk("RESUELVE:\n", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Reconocer y pagar "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ")+listaDatosDocentes.get(0).get("TIPO_VINCULACION")+" ");
        c = new Chunk("" + NOM_DOCENTE + ",", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
       
        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? " identificado con " : " identificada con "));
        if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk("" + getNumeroDecimal(("" + totalpuntos).replace(".", ",")) + " (" + ValidarNumeroDec("" + totalpuntos) + ") puntos salariales, ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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

        Cell celdaProductos = new Cell(new Paragraph("N??", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Producto", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);
        celdaProductos = new Cell(new Paragraph("Nombre del Producto", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("N?? Acta", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Puntos Reconocidos", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        if (listaDatosDocentesCargoAcad.size() == 0) {
            celdaProductos = new Cell(new Paragraph("Fecha a partir de la cual surten efectos fiscales", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
        }

        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            celdaProductos = new Cell(new Paragraph("" + (i + 1), Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("TIPO_PRODUCTO")).equals("Art_Col") ? "Art??culo" : listaDatosDocentes.get(i).get("TIPO_PRODUCTO").replace("_", " ")), Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
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
            celdaProductos = new Cell(new Paragraph("" + nombresolicitud, Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + listaDatosDocentes.get(i).get("ACTA"), Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            
            try{
            celdaProductos = new Cell(new Paragraph("" + getNumeroDecimal(listaDatosDocentes.get(i).get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get(i).get("PUNTOS")) + ") puntos", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
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
                    celdaProductos = new Cell(new Paragraph("" + fechaEnletras(listaDatosDocentes.get(i).get("RETROACTIVIDAD"),0), Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
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
                    celdaProductos = new Cell(new Paragraph("" , Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("\n");
        c = new Chunk("Par??grafo ??nico: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Los puntos salariales asignados en el recuadro anterior, se establecen seg??n lo dispuesto en el "
                + "aparte de ???Decisi??n??? ");
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
            c = new Chunk(" del Acta N?? " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
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

        p.add(" del Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje respectivamente. ");
        if(listaDatosDocentesCargoAcad.size()>0)
            p.add("As?? mismo, surtir??n efectos fiscales una vez "+(listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el " : "la ")+"docente finalice la comisi??n para desempe??ar el cargo de libre nombramiento y remoci??n que le ha sido otorgado.");
        else    
            p.add("As?? mismo, tendr??n efectos fiscales a partir de la fecha en que el Comit?? expidi?? "
                + "el acto formal de reconocimiento, seg??n consta en la relaci??n anterior.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n "
                + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk("" + Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto, env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n"
                + ", haci??ndole saber que contra la misma procede recurso de reposici??n ante el mismo funcionario "
                + "que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as siguientes "
                + "a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de La Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales, "
                + "a fin de que procedan a realizar el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno "
                + "de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida ");

        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk("" + Utilidades.Utilidades.decodificarElemento(NOM_DOCENTE) + ".", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("La presente resoluci??n rige a partir del t??rmino de su ejecutoria.");
        p.add("\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setFont(Fonts.SetFont(Color.black, 11, Fonts.NORMAL));
        p.setAlignment(centrado);
        c = new Chunk("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);

        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Dada en la ciudad de Santa Marta, D. T. C. H., a los");
        p.add("\n \n \n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("PABLO VERA SALAZAR");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Rector ");
        p.add("\n\n");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add("" + listaProyecto.get("PROYECTO") + "____");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add("" + listaProyecto.get("REVISO") + "____");
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(left);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
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
        Fonts f = new Fonts(arialFont);
        
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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        String resolucion = "\"Por la cual se promociona en el Escalaf??n Docente a la categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))
                            + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                            + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        celda = new Cell(new Paragraph(resolucion+"", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
       
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el Art??culo 20 del Acuerdo Superior N?? 007 de 2003 establece los criterios para la clasificaci??n en el escalaf??n"
                + " del personal docente de la Universidad del Magdalena, aplicables a quienes se encuentren amparados por el Decreto N?? 1279 de 2002, o las normas que lo modifiquen o sustituyan."
                + "\n");
                documento.add(p);
                
                
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el art??culo 24 del citado Acuerdo, se??ala los requerimientos para promover a un profesor en la carrera docente."
                + "\n");
        documento.add(p);

        if(!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("AS_ANTIGUO")).equals("N/A")){

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que por su parte, el art??culo 27 ibidem se??ala los requisitos que deben cumplir los profesores nombrados en planta para ser promovidos en el escalaf??n"
                    + ", indicando para la categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " los siguientes: \n");
            documento.add(p);

            String requisitosxCategoria = "";
            if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asistente")) {
                requisitosxCategoria = "a. Ser profesor Auxiliar de tiempo completo o de medio tiempo y acreditar dos (2) a??os de permanencia en dicha categor??a.\n"
                        + "b. Haber sido evaluado satisfactoriamente en el desempe??o de las funciones durante los dos (2) ??ltimos procesos de evaluaci??n docente integral que semestralmente practica la Universidad.\n"
                        + "c. Acreditar m??nimo cuarenta (40) horas de formaci??n pedag??gica. \n"
                        + "d. Presentar T??tulo de Especializaci??n en el ??rea de titulaci??n Profesional.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asociado")) {
                requisitosxCategoria = "a. Ser profesor Asistente de tiempo completo o de medio tiempo y acreditar "
                        + "dos (2) a??os de permanencia en dicha categor??a.\n"
                        + "b. Haber sido evaluado satisfactoriamente en el desempe??o de sus funciones "
                        + "durante los dos (2) ??ltimos procesos de evaluaci??n docente integral que "
                        + "semestralmente practica la Universidad.\n"
                        + "c. Tener t??tulo de Maestr??a.\n"
                        + "d. Someter y obtener la aprobaci??n de pares externos de un (1) trabajo realizado "
                        + "durante su permanencia en la categor??a de profesor auxiliar, el cual constituya "
                        + "un aporte significativo a la docencia, a las ciencias, a las artes o a las "
                        + "humanidades de acuerdo con la reglamentaci??n que para el efecto expida el "
                        + "Consejo Superior.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Titular")) {
                requisitosxCategoria = "a. Cumplir los requisitos exigidos para ser profesor asociado.\n"
                        + "b. Acreditar tres (3) a??os de experiencia calificada en el equivalente de tiempo "
                        + "completo como profesor Asociado.\n"
                        + "c. Haber sido evaluado satisfactoriamente en el desempe??o de sus funciones "
                        + "durante los dos ??ltimos procesos de evaluaci??n docente integral que "
                        + "semestralmente practica la Universidad.\n"
                        + "d. Tener t??tulo de doctorado.\n"
                        + "e. Someter y obtener la aprobaci??n de pares externos de dos (2) trabajos "
                        + "realizados durante su permanencia en la categor??a de profesor asociado que "
                        + "constituyan un aporte significativo a la docencia, a las ciencias, a las artes o "
                        + "a las humanidades de acuerdo con la reglamentaci??n que para el efecto "
                        + "expida el Consejo Superior.\"";
            }

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(Fonts.SetFont(Color.black, 8, Fonts.ITALIC));
            c = new Chunk("\"... ART??CULO 27. ",Fonts.SetFont(Color.black, 8, Fonts.BOLD) );
            p.add(c);
            p.add("Los profesores nombrados en planta seg??n la categor??a deben cumplir los siguientes requisitos: \n"
            +"... \n");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+"\n",Fonts.SetFont(Color.black, 8, Fonts.BOLD));
            p.add(c);
            p.add( requisitosxCategoria + "\n");
            documento.add(p);
        }
        else{
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el Art??culo 14 del Acuerdo Superior N?? 007 de 2021 establece los requisitos m??nimos que debe cumplir un profesor de planta"
                + " para ser clasificado al momento de su vinculaci??n en una de las categor??as establecidas por el Decreto 1279 de 2002,"
                + " determinando para la categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+" lo siguiente: \n");
                documento.add(p);
            String requisitosxCategoria = "";
            if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asistente")) {
                requisitosxCategoria = "??? T??tulo de posgrado.\n" +
                          "??? Cinco (5) a??os de experiencia certificada, en tiempo completo equivalente, en actividades profesorales en programas acad??micos "
                        + "de educaci??n universitaria o en actividades profesionales afines al perfil en el que va a ser nombrado\n" +
                          "??? Productividad acad??mica equivalente a veinte (20) puntos salariales.\"\n";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Asociado")) {
                requisitosxCategoria = "??? T??tulo de maestr??a o doctorado.\n" +
                          "??? Siete (7) a??os de experiencia certificada, en tiempo completo equivalente, "
                        + "en actividades profesorales en programas acad??micos "
                        + "de educaci??n universitaria o en actividades profesionales afines al perfil en el que va a ser nombrado.\n" +
                          "??? Productividad acad??mica equivalente a cuarenta (40) puntos salariales.\"";
            } else if (Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")).equals("Profesor Titular")) {
                requisitosxCategoria = "??? T??tulo de doctorado.\n" +
                          "??? Acreditar tres (3) a??os de experiencia calificada en el equivalente a tiempo completo como profesor asociado"
                        + " de la Universidad del Magdalena.\n" +
                          "??? Productividad acad??mica equivalente a ochenta (80) puntos salariales.\"";
            }
            
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(Fonts.SetFont(Color.black, 8, Fonts.ITALIC));
            p.add("\"");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD"))+"\n \n",Fonts.SetFont(Color.black, 8, Fonts.BOLD));
            p.add(c);
            p.add( requisitosxCategoria + "\n");
            documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que la categor??a dentro del escalaf??n docente es uno de los factores que incide en la modificaci??n de puntos salariales de los docentes de planta,"
                + " de acuerdo con lo establecido en el Literal b. del art??culo 12 del Decreto 1279 de 2002.\n"
                + "\n"
                + "Que el art??culo 8?? de la precitada norma, establece la asignaci??n de puntaje por categor??a acad??mica en el escalaf??n para docentes en carrera, cualquiera que sea su dedicaci??n.\n");
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
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Art??culo 17 del referido decreto, establece los criterios a tener en cuenta para la modificaci??n de salario "+
                    "de los docentes que realizan actividades acad??mico-administrativas, disponi??ndose, en igual sentido, en el art??culo 62 de la mencionada "+
                    "disposici??n, que el Grupo de Seguimiento al r??gimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la informaci??n a nivel nacional, y adem??s adecuar los criterios "+
                    "y efectuar los ajustes a las metodolog??as de evaluaci??n aplicadas por los Comit??s Internos de Asignaci??n de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al R??gimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N?? 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad acad??mica para los docentes que asuman cargos acad??mico-administrativos, este "+
                    "se??al?? que independientemente de haber elegido entre la remuneraci??n del cargo que va a desempe??ar y la que le corresponde como docente, "+
                    "solo se podr?? hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades acad??mico-administrativas, de conformidad con el Art??culo 17 del Decreto 1279 de 2002 y el par??grafo 1 del Art??culo 60 del Acuerdo N?? 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que mediante Resoluci??n Rectoral N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add("fue comisionado para ejercer un cargo de libre nombramiento y remoci??n dentro de la Instituci??n como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tom?? posesi??n mediante Acta N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }
        
        if(!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("AS_ANTIGUO")).equals("N/A")){
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("solicit?? promoci??n en el escalaf??n docente de la categor??a " + categoriaAnterior + " a la categor??a " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " de conformidad con lo prescrito por el Art??culo 22 del Acuerdo Superior N?? 007 de 2003.\n");
        documento.add(p);

       } 
        else{
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("solicit?? promoci??n en el escalaf??n docente de la categor??a " + categoriaAnterior + " a la categor??a " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + " de conformidad con lo prescrito por el Art??culo 14 del Acuerdo Superior N?? 007 de 2021.\n");
        documento.add(p);
        }
        
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        try{
        p.add("Que el Consejo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FACULTAD")) + ", en sesi??n llevada a cabo el d??a " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0)
                + " contenida en Acta N?? " + getNumeroDecimal(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_CONS_FACULTAD"))) + ", estudi?? la solicitud y emiti?? concepto favorable de ascenso en el escalaf??n docente ante Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje.\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        try{
        p.add("Que mediante Acta N?? " + listaDatosDocentes.get("ACTA") + " de " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ", tal y como quedo establecido en el ??tem " + listaDatosDocentes.get("NUMERAL_ACTA_CIARP")
                +", el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje decidi?? promover en el escalaf??n docente "+ (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al profesor" : "a la profesora"));
        
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
        c = new Chunk(" " +Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("a la categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ", por cumplir con los requisitos fijados para ello y, asign?? ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        p.add(" de acuerdo con el par??grafo II del art??culo 8 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que conforme lo dispuesto en el art??culo 31 del Acuerdo Superior N?? 007 de 2003,"
                + " el Rector expedir?? la Resoluci??n de ascenso, previa recomendaci??n motivada del Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje.\n"
                + "\n"
                + "Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al a??o el total de puntos que corresponde a cada"
                + " docente, conforme lo se??ala el Art??culo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Promover en el escalaf??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        p.add("a la Categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Autorizar el reconocimiento y pago de ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("correspondientes a su promoci??n de Categor??a de " + categoriaAnterior + " a Categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        
        p.add("Los puntos salariales");
                
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p.add(" surtir??n efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisi??n para desempe??ar el cargo de libre nombramiento y remoci??n que le ha sido otorgado.\n");
        }else{
            if(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RETROACTIVIDAD")).length()==10){
                try{
            p.add(" surtir??n efectos fiscales a partir del d??a "
                + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RETROACTIVIDAD")),0) + ", fecha en la cual se expidi?? el acto formal de "+
                "reconocimiento seg??n consta en Acta N?? " + listaDatosDocentes.get("ACTA") + ".\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO QUINTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEXTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
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
        Fonts f = new Fonts(arialFont);

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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconoc"
                + "imiento y pago de bonificaci??n por productividad acad??mica"
                + " " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>

        celda = new Cell(new Paragraph(resolucion+ "", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
      
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
       
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el Art??culo 19 del Decreto N?? 1279 de 2002, consagra las bonificaciones como reconocimientos monetarios no salariales, que se reconocen por una sola vez, correspondientes a actividades espec??ficas de productividad acad??mica y no contemplan pagos gen??ricos indiscriminados.\n"
                + "\n"
                + "Que en el Art??culo 20 de la disposici??n anterior, se establecieron los criterios para el reconocimiento de bonificaciones productividad acad??mica, siendo determinados en igual sentido y en consonancia con esta disposici??n, en los Art??culos 51 y 52 del Acuerdo Superior N?? 007 de 2003.\n"
                + "\n"
                + "Que conforme lo establecido en los incisos tercero y cuarto del Art??culo 19 del Decreto N?? 1279 de 2002, las universidades deben reconocer y pagar semestralmente las bonificaciones que se causen en dichos periodos, de igual forma, las actividades de productividad acad??mica que tengan reconocimiento salarial no reciben bonificaciones.\n");
        documento.add(p);
        
        if(listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE TIEMPO COMPLETO") || 
            listaDatosDocentes.get(0).get("TIPO_VINCULACION").toUpperCase().equals("DOCENTE OCASIONAL DE MEDIO TIEMPO")    ){
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que el Art??culo 71 del Acuerdo Superior N?? 007 de 19 de marzo de 2003, establece que los Docentes Ocasionales participan de las bonificaciones por productividad acad??mica (Cap??tulo VI del Decreto 1279 de 19 de junio de 2002), por actividades realizadas durante el periodo en que tienen vinculaci??n.\n");
            documento.add(p);
        }

       
       
        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            totalPuntos += Double.parseDouble(listaDatosDocentes.get(i).get("PUNTOS").replace(",", "."));
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);

        p.add(", present?? solicitud de asignaci??n de puntos de bonificaci??n por ");

        c = new Chunk((listaDatosDocentes.size()==1 ? "un " :numeroEnLetras(listaDatosDocentes.size())) + " (" + listaDatosDocentes.size() + ")", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
            c = new Chunk(" del Acta N?? " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
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
                    numerales += (numerales.equals("")?"":", ")+"??tem "+ 
                            
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
                    numerales += (numerales.equals("")?"":", ")+"??tem "+ 
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
                        norma +=(norma.equals("")?"":", ")+ "??tem 1 literal h. numeral II, Art??culo 20 del Decreto 1279 del 2002";
                    } else if (tipoTesis[0].equalsIgnoreCase("Tesis de Doctorado")) {
                        norma += (norma.equals("")?"":", ")+ "??tem 2 literal h. numeral II, Art??culo 20 del Decreto 1279 del 2002";
                    }
                    
                } else {
                    numerales += (numerales.equals("")?"":", ")+"??tem "+ 
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
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            try{
            p.add("Que en virtud de lo anterior y teniendo en cuenta lo se??alado en el " + norma + " el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje,"
                    + " en sesi??n realizada el d??a " + fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0) + " contenida en el Acta N?? " + listaActas.get(j).get("ACTA")
                    + " asign?? ");
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
            c = new Chunk(getNumeroDecimal(""+totalPuntosxActas) + " (" + ValidarNumeroDec(""+totalPuntosxActas) + ") puntos de bonificaci??n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
                
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
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE"))+",", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add(" distribuidos en " + (listaDatosDocentes.size() > 1 ? "los" : "el") + " " + (listaDatosDocentes.size() > 1 ? "productos presentados" : "producto presentado")
                    + ", seg??n el orden establecido en el apartado \"Decisi??n\" del ");
           try{
            p.add("" +numerales
                +" del Acta N?? "+listaActas.get(j).get("ACTA")+" de "+fechaEnletras(listaActas.get(j).get("FECHA_ACTA"),0)+".\n");
            
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al a??o el total de puntos"
                + " que corresponde a cada docente, conforme lo se??ala el Art??culo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add((listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get(0).get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get(0).get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(""+totalPuntos) + " (" + ValidarNumeroDec(""+totalPuntos) + ") puntos de bonificaci??n", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            
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
        p.add(" por la productividad acad??mica que se relaciona en la siguiente tabla:\n");
        documento.add(p);

        ///TABLA
        //<editor-fold defaultstate="collapsed" desc="TABLA">
        float[] tamy = new float[]{5f, 13f, 50f, 8f, 13f, 20f};
        Table TableProductos = new Table(6);
        TableProductos.setWidths(tamy);
        TableProductos.setWidth(100);

        Cell celdaProductos = new Cell(new Paragraph("N??", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Producto", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);
        celdaProductos = new Cell(new Paragraph("Nombre del Producto", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("N?? Acta", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Fecha de Acta", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        celdaProductos = new Cell(new Paragraph("Puntos Reconocidos", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
        celdaProductos.setBorder(15);
        celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
        TableProductos.addCell(celdaProductos);

        for (int i = 0; i < listaDatosDocentes.size(); i++) {
            celdaProductos = new Cell(new Paragraph("" + (i + 1), Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(i).get("TIPO_PRODUCTO")).replace("_", " "), Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
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
            celdaProductos = new Cell(new Paragraph("" + nombresolicitud, Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + listaDatosDocentes.get(i).get("ACTA"), Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);

            celdaProductos = new Cell(new Paragraph("" + fechaEnletras(listaDatosDocentes.get(i).get("FECHA_ACTA"),0), Fonts.SetFont(Color.black, 7, Fonts.NORMAL)));
            celdaProductos.setBorder(15);
            celdaProductos.setHorizontalAlignment(Cell.ALIGN_CENTER);
            celdaProductos.setVerticalAlignment(Cell.ALIGN_CENTER);
            TableProductos.addCell(celdaProductos);
            
            try{
            celdaProductos = new Cell(new Paragraph("" + getNumeroDecimal(listaDatosDocentes.get(i).get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get(i).get("PUNTOS")) + ") puntos", Fonts.SetFont(Color.black, 7, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("\nPar??grafo:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
         
        p.add(" Los ");
        try{
        c = new Chunk(getNumeroDecimal(""+totalPuntos) + " (" + ValidarNumeroDec(""+totalPuntos) + ") puntos de bonificaci??n", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        p.add(" por productividad acad??mica asignados en el recuadro anterior, se establecen seg??n lo dispuesto en el aparte de \"Decisi??n\" ");

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
            c = new Chunk(" del Acta N?? " + listaActas.get(i).get("ACTA") + " de " + fechaEnletras(listaActas.get(i).get("FECHA_ACTA"),0) + "", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            
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
        
        
        p.add(" los cuales, se reconocer??n y pagar??n por una sola vez" + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? " al" : " a la") + " "+ listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", de conformidad con lo dispuesto en el ART??CULO 19 del Decreto 1279 de 19 de junio de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get(0).get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get(0).get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get(0).get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        
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
        Fonts f = new Fonts(arialFont);
        
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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconocimiento y pago de puntos salariales por el t??tulo de " + tituloCorto[0]
                            + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                            + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        
        celda = new Cell(new Paragraph(resolucion + "", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
      
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
      
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el art??culo 62 del Decreto 1279 de junio de 2002, establece que el Grupo de Seguimiento al r??gimen salarial y prestacional de los profesores universitarios,"
                + " puede definir las directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la informaci??n a nivel nacional, y adem??s adecuar los criterios"
                + " y efectuar los ajustes a las metodolog??as de evaluaci??n aplicadas por los Comit??s Internos de Asignaci??n de Puntaje o los organismos que hagan sus veces.\n"
                + " \n"
                + "Que el art??culo primero, numeral 22 del Acuerdo No. 001 de 4 de marzo de 2004, del Grupo de Seguimiento del r??gimen salarial y prestacional de los profesores universitarios"
                + " del Decreto 1279 de 19 de junio de 2002, se??ala sobre la asignaci??n de puntaje por t??tulos acad??micos obtenidos en el exterior lo siguiente:\n");
        documento.add(p);
        
        

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setIndentationLeft(13);
        p.setIndentationRight(12);
        p.setFont(Fonts.SetFont(Color.black, 8, Fonts.ITALIC));
        p.add("\"22. El Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje debe dar tr??mite a las solicitudes en la medida en que se vayan presentando; las modificaciones salariales tendr??n efectos a partir de la fecha en que el Comit?? expida el acto formal de reconocimiento. \n"
                +"???\n"
                + "En cuanto la asignaci??n de puntaje por t??tulos acad??micos obtenidos en el exterior es procedente aplicar lo dispuesto en el Art??culo 12 del Decreto 861 de 2000, el cual se??ala"
                + " ???Art??culo 12. De los t??tulos y certificados obtenidos en el exterior requerir??n para su validez, de las autenticaciones, registros y equivalencias determinadas por el Ministerio de Educaci??n Nacional"
                + " y el Instituto Colombiano para el Fomento de la Educaci??n Superior.\n"
                + "???\n"
                + "De otro lado si el docente aporta el t??tulo para modificaci??n de salario y no cumple con el requisito de la convalidaci??n dentro del plazo se??alado, el rector mediante acto administrativo determinar?? "
                + "que no reconoce los puntos asignados por el incumplimiento de la condici??n se??alada en el acto expedido por el CIARP (Art??culo 55 Decreto 1279 de 2002).\n"
                + "\n"
                + "De presentarse la convalidaci??n en el t??rmino se??alado, en la resoluci??n rectoral que se debe expedir dos veces al a??o (Art??culo 55 Decreto 1279 de 2002) se ordenar?? el pago del salario modificado, "
                + "desde el momento en que se reconoci?? y asign?? el respectivo puntaje por parte del CIARP, de acuerdo a lo dispuesto en el par??grafo III del Art??culo 12 del Decreto 1279,"
                + " el cual dispone: ???las modificaciones salariales tienen efecto a partir de la fecha en que el Comit?? Interno de Asignaci??n y Reconocimiento del Puntaje, "
                + "o el ??rgano que haga sus veces en cada una de las universidades, expida el acto formal de reconocimiento de los puntos salariales asignados en el marco del presente decreto??? ???.\"\n");
        documento.add(p);
        
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Art??culo 17 del referido decreto, establece los criterios a tener en cuenta para la modificaci??n de salario "+
                    "de los docentes que realizan actividades acad??mico-administrativas, disponi??ndose, en igual sentido, en el Art??culo 62 de la mencionada "+
                    "disposici??n, que el Grupo de Seguimiento al r??gimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la informaci??n a nivel nacional, y adem??s adecuar los criterios "+
                    "y efectuar los ajustes a las metodolog??as de evaluaci??n aplicadas por los Comit??s Internos de Asignaci??n de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al R??gimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N?? 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad acad??mica para los docentes que asuman cargos acad??mico-administrativos, este "+
                    "se??al?? que independientemente de haber elegido entre la remuneraci??n del cargo que va a desempe??ar y la que le corresponde como docente, "+
                    "solo se podr?? hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades acad??mico-administrativas, de conformidad con el Art??culo 17 del Decreto 1279 de 2002 y el Art??culo 6 del Acuerdo N?? 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que mediante Resoluci??n Rectoral N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", "+
                    "fue comisionado para ejercer un cargo de libre nombramiento y remoci??n dentro de la Instituci??n como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tom?? posesi??n mediante Acta N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 8, Fonts.BOLD));
        p.add(c);
        p.add("realiz?? solicitud de asignaci??n de puntos salariales por el t??tulo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        try{
        p.add("Que el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, en sesi??n realizada el d??a " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA_CIARP_INICIAL"),0)
                + " contenida en Acta N?? " + listaDatosDocentes.get("ACTA") + ", asign?? ");
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
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + getNumero(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        try{
        p.add("por el t??tulo de " + listaDatosDocentes.get("TITULO") + ", que se pagar??n con efectos a partir del "+ fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0)+" al convalidar el t??tulo, para lo cual, el docente cuenta hasta el "+ fechaEnletras(listaDatosDocentes.get("FECHA_MAX_CONV"),0)+ ".\n");
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
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que mediante resoluci??n rectoral N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RES_TIT_ANT")) + " se reconoce " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al " : "a la ") + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
           try{
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + " " + getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
            p.add(" correspondientes al t??tulo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add("Que en dicho acto administrativo se estableci?? que para la asignaci??n y pago de los puntos, el docente debe cumplir con la convalidaci??n del t??tulo dentro de los dos (2) a??os siguientes, "
                    + "contados a partir del d??a " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", fecha en la que el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje asign?? los puntos salariales.\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el " : "la ") + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" present?? ante el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, copia de la " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        try{
        p.add("Que el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, en sesi??n realizada el d??a " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                + " contenida en Acta N?? " + listaDatosDocentes.get("ACTA") + " punto "+ listaDatosDocentes.get("NUMERAL_ACTA_CIARP")+", determin?? tener por cumplido el requisito de la convalidaci??n del t??tulo,"
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
        c = new Chunk(numeroEnLetras(Integer.parseInt(listaDatosDocentes.get("PUNTOS"))) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
            p.add(" surtir??n efectos fiscales desde el " );
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
            p.add(" surtir??n efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisi??n a la que fue "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "encargado" : "encargada")+".\n");
        }
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al a??o el total de puntos que corresponde a cada"
                + " docente, conforme lo se??ala el Art??culo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        
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
        p.add(" por el t??tulo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Los ");
        try{
        c = new Chunk(numeroEnLetras(Integer.parseInt(listaDatosDocentes.get("PUNTOS"))) + " (" + listaDatosDocentes.get("PUNTOS") + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
            p.add(" surtir??n efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisi??n para desempe??ar el cargo de libre nombramiento y remoci??n que le ha sido otorgado.");
        }else{
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add(" surtir??n efectos fiscales a partir del d??a " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", seg??n consta en Acta N?? " + listaDatosDocentes.get("ACTA") + " del " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ".\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO QUINTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
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
        Fonts f = new Fonts(arialFont);
        
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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        celda = new Cell(new Paragraph(resolucion+"", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que la Ley 30 de 1992 en su Art??culo 28 reconoce a las Universidades en virtud del principio de autonom??a universitaria,"
                + " entre otros, el derecho a seleccionar a sus profesores.\n");
        documento.add(p);

        if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FA_BECA")).equals("N/A")) {

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que mediante Acuerdo Superior N?? 025 de 2002, modificado por el Acuerdo Superior N?? 008 de 2014, se adopt?? el Programa de Formaci??n Avanzada para la Docencia y la Investigaci??n, "
                    + "disponiendo en su Art??culo S??ptimo la facultad del Rector de la Universidad del Magdalena para vincular a la planta de personal docente,"
                    + " a los beneficiarios de becas concedidas por organismos o entidades de reconocido prestigio nacional o internacional diferentes a esta instituci??n.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el se??or " : "la se??ora "));
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
            if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
                p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + listaDatosDocentes.get("C_EXPEDICION") + ", ");
            } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
                p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            } else {
                p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            }
            p.add("result?? " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "beneficiario " : "beneficiaria ") + "de un(a) " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FA_BECA"))
                    + ", siendo por ello, vinculado de manera excepcional a la Universidad como Docente de Planta a trav??s del Programa de Formaci??n Avanzada para la Docencia y la Investigaci??n,"
                    + " de conformidad con el Acuerdo Acad??mico N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACUERDO_FA"))+".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que en ese orden, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el se??or " : "la se??ora "));
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add(" fue " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "nombrado " : "nombrada ") + "como " + listaDatosDocentes.get("TIPO_VINCULACION")
                    + " mediante Resoluci??n Rectoral N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_INGRESO")) + ", cargo del cual tom?? posesi??n mediante Acta N?? "
                    + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_INGRESO")) + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que luego de haber finalizado los estudios, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " docente en menci??n fue reincorporado al servicio como " + listaDatosDocentes.get("TIPO_VINCULACION")
                    + " de la Universidad a trav??s de la Resoluci??n Rectoral N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_REINTEGRO")) + ".\n");
            documento.add(p);
        } else if (!Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_INICIO")).equals("N/A")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que mediante Resoluci??n Rectoral No. " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_INICIO")) + ", se dio inicio a la convocatoria p??blica para proveer cargos docentes en dedicaci??n de Tiempo Completo.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que seg??n Resoluci??n Rectoral N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("INFO_CONCURSO_FIN")));
            c = new Chunk(" " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
            if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
                p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + listaDatosDocentes.get("C_EXPEDICION") + ", ");
            } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
                p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            } else {
                p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
            }
            p.add("result?? " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "ganador" : "ganadora") + " del concurso en el ??rea de desempe??o " + listaDatosDocentes.get("AREA_DESEMPE??O")
                    + " " + listaDatosDocentes.get("FACULTAD") + ".\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que mediante Resoluci??n Rectoral N?? " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_INGRESO") )+ " fue " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "nombrado " : "nombrada ")
                    + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el se??or " : "la se??ora "));

            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add(" en el cargo de " + listaDatosDocentes.get("TIPO_VINCULACION") + " de la Universidad del Magdalena, cargo del que tom?? posesi??n mediante Acta N?? "
                    + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_INGRESO")) + ".\n");
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el Acuerdo Superior N?? 007 de 2003 establece en su Art??culo 14, que los profesores aspirantes a la carrera docente se nombrar??n de tiempo completo o medio tiempo por un periodo de prueba"
                + " de un (1) a??o; vencido y aprobado el mismo, el profesor podr?? solicitar su ingreso a la carrera ante el Comit?? de Asignaci??n y Reconocimiento de Puntaje en la categor??a que le corresponda en el escalaf??n.\n"
                + "\n"
                + "Que mediante el Art??culo 24 del acuerdo en menci??n, se establecieron los requisitos para promover a un profesor en la carrera docente, se??al??ndose en igual sentido en el Par??grafo de esta disposici??n, "
                + "que los profesores con t??tulo universitario que cumplan y aprueben el periodo de prueba, ingresan al escalaf??n docente en la categor??a de profesor auxiliar y en la categor??a correspondiente para docentes sin t??tulo universitario.\n"
                + "\n"
                + "Que por otra parte, el Par??grafo 1?? del Art??culo 27 del Acuerdo Superior N?? 007 de 2003, establece como excepci??n de la norma antes citada, que el profesor vinculado, escalafonado previamente"
                + " en una universidad p??blica con un sistema similar al de la Universidad del Magdalena, ser?? ubicado en su categor??a, despu??s de superar el periodo de prueba, previa constancia expedida por la universidad de procedencia.\n");
        documento.add(p);

        if (listaDatosDocentes.get("PENDIENTE_INGLES").equalsIgnoreCase("SI")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que mediante Acuerdo Superior N?? 008 del 19 de mayo de 2008, se reglament?? el requisito de certificaci??n de la prueba de suficiencia en el manejo del ingl??s para la provisi??n de cargos docentes en la Universidad.\n"
                    + "\n"
                    + "Que a trav??s del Acuerdo Superior N?? 009 del 03 de mayo de 2013, se modific?? el Art??culo 1 del Acuerdo Superior N?? 008 de 2008, reglament??ndose como m??nimo un Nivel B2 seg??n el Marco Com??n de Referencia Europeo"
                    + " para acreditar la suficiencia en el manejo del idioma ingl??s mediante ex??menes con validez internacional o estudios desarrollados en pa??ses de habla inglesa.\n"
                    + "\n"
                    + "Que el mencionado acuerdo dispone que el aspirante vinculado a la Universidad cuenta con un m??ximo de diez (10) meses contados a partir de su vinculaci??n para"
                    + " acreditar el manejo en el idioma ingl??s so pena que la evaluaci??n de su desempe??o en el a??o de prueba correspondiente sea declarada no satisfactoria.\n"
                    + "\n"
                    + "Que el Acuerdo Superior N?? 015 del 30 de noviembre de 2016 modific?? el Par??grafo Segundo del art??culo Primero del Acuerdo Superior N?? 009 de 2013 quedando as??:\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setIndentationLeft(13);
            p.setIndentationRight(12);
            p.setFont(Fonts.SetFont(Color.black, 8, Fonts.NORMAL));
            p.add("\"Par??grafo 2: El aspirante que no certifique el m??nimo de suficiencia en el manejo del idioma ingl??s, podr?? ser incluido en la lista de elegibles siempre y cuando haya obtenido "
                    + "el puntaje m??nimo requerido del total de la calificaci??n final en el proceso de selecci??n. Si el aspirante es vinculado a la universidad, contar?? con un m??ximo de hasta veintid??s (22)"
                    + " meses contados a partir de su vinculaci??n para acreditar el manejo del idioma ingl??s so pena de impedirse su ascenso a la categor??a de profesor asistente en el escalaf??n docente\"\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            try{
            p.add("Que el Consejo de " + listaDatosDocentes.get("FACULTAD") + " en sesi??n celebrada el d??a " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0) + " contenida en Acta N?? "
                    + listaDatosDocentes.get("ACTA_CONS_FACULTAD") + ", determin?? que la evaluaci??n del per??odo de prueba " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
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
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add("fue superada de manera satisfactoria, realizando la claridad que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " docente debe cumplir con el requisito de manejo del idioma ingl??s.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            try{
            p.add("solicit?? ingreso a la carrera docente ante el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, ??rgano que en sesi??n realizada el d??a " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                    + " contenida en Acta N?? " + listaDatosDocentes.get("ACTA") + ", verific?? el cumplimiento de los requisitos establecidos para ??ste y la superaci??n del periodo de prueba, aclarando que tiene pendiente acreditar el manejo del idioma ingl??s.\n");
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
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            try{
            p.add("Que el Consejo de " + listaDatosDocentes.get("FACULTAD") + " en sesi??n celebrada el d??a " + fechaEnletras(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("FECHA_ACTA_CF")), 0) + " contenida en Acta N?? "
                    + listaDatosDocentes.get("ACTA_CONS_FACULTAD") + " , determin?? que la evaluaci??n del per??odo de prueba " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " "
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
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add("fue superada de manera satisfactoria.\n");
            documento.add(p);

            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                    + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            try{
            p.add("solicit?? ingreso a la carrera docente ante el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, ??rgano que en sesi??n realizada el d??a " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                    + " contenida en Acta N?? " + listaDatosDocentes.get("ACTA") + ", verific?? el cumplimiento de los requisitos establecidos para ??ste y la superaci??n del periodo de prueba.\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Autorizar el ingreso a la carrera " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        p.add("en la categor??a de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CAT_INGRESO")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("El ingreso a la carrera "+ (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
            try{
        p.add(", producir??n efectos fiscales a partir del " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", fecha en la que el Comit?? verific?? el cumplimiento de los requisitos, seg??n consta en Acta N?? " + listaDatosDocentes.get("ACTA") + " del " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0) + ".\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para acreditar el manejo del idioma ingl??s, " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("tiene un plazo de veintid??s (22) meses contados a partir de la fecha de reincorporaci??n que se produjo mediante Resoluci??n Rectoral N?? "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_REINTEGRO"))
        +", so pena de impedirse su ascenso a la categor??a de profesor asistente, conforme lo se??ala el art??culo 1 del Acuerdo Superior N?? 015 de 2016.\n");
        documento.add(p);
           
            p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO QUINTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEXTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);
       }else{
        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ". ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + "el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO QUINTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);
        }
       
        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 8, Fonts.BOLD));
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
        Fonts f = new Fonts(arialFont);
       
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
        celda = new Cell(new Paragraph("DESPACHO DEL RECTOR\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        celda = new Cell(new Paragraph("RESOLUCI??N N??\n", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        headerTable.addCell(celda);
        
        String resolucion = "\"Por la cual se autoriza el reconocimiento y pago de puntos salariales por el t??tulo de " + tituloCorto[0]
                + " " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la")
                + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + "\"";
        
        //<editor-fold defaultstate="collapsed" desc="Datos para Excel">
            datos1 = new HashMap<>();
            datos1.put("No", "1");
            datos1.put("RESOLUCION", ""+resolucion);
            datos2.add(datos1);
        //</editor-fold>
        
        celda = new Cell(new Paragraph(resolucion+"", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
       

        celda = new Cell(new Paragraph("\nUNIVERSIDAD DEL MAGDALENA - RECTOR??A ??? Resoluci??n N?? ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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

        celda = new Cell(new Paragraph("P??gina ", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_RIGHT);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);
        celda = new Cell(new RtfPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        
        footertable.addCell(celda);

        celda = new Cell(new Paragraph("de", Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
        celda.setBorder(0);
        celda.setHorizontalAlignment(Cell.ALIGN_CENTER);
        celda.setVerticalAlignment(Cell.ALIGN_BOTTOM);
        
        footertable.addCell(celda);

        celda = new Cell(new RtfTotalPageNumber(Fonts.SetFont(Color.black, 8, Fonts.BOLD)));
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("El Rector de la Universidad del Magdalena ");
        Chunk c = new Chunk("???UNIMAGDALENA???", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(", en ejercicio de sus funciones legales y en especial las que le confiere el Estatuto General, el Estatuto Docente de la Universidad, y\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("CONSIDERANDO:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el Cap??tulo III del Decreto 1279 de 19 de junio de 2002, establece los factores y criterios que inciden en la modificaci??n de puntos salariales "
                + "de los docentes amparados por este r??gimen.\n"
                + "\n"
                + "Que los t??tulos correspondientes a estudios universitarios de pregrado o posgrado es uno de los factores que incide en las modificaciones de los puntos"
                + " salariales de los docentes de planta, de acuerdo con lo establecido en el Literal a. del art??culo 12 del Decreto 1279 del 19 de junio de 2002.\n");
        documento.add(p);

        String titulo = "";
        String literal ="";
        
        if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Doctorado")) {
            titulo = "PhD. o Doctorado";
            literal ="art??culo 7 Numeral 2, Literal c del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Maestr??a")) {
            titulo = "Magister o Maestr??a";
            literal ="art??culo 7 Numeral 2, Literal b del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Especializaci??n")) {
            titulo = "Especializaci??n";
            literal ="art??culo 7 Numeral 2, Literal a del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Especializaci??n Cl??nica")) {
            titulo = "Especializaci??n Cl??nica";
            literal ="art??culo 7 Numeral 2, par??grafo II del Decreto 1279 del 19 de junio de 2002,";
        } else if (listaDatosDocentes.get("SUBTIPO_PRODUCTO").equals("Pregrado")) {
            titulo = "Pregrado";
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que el "+literal+" establece la asignaci??n de puntaje por t??tulos universitarios de posgrado,"
                + " disponiendo para el t??tulo de " + titulo + " lo siguiente: \n");
        documento.add(p);

        String requisitosxTitulo = "";
        if (titulo.equals("PhD. o Doctorado")) {
            requisitosxTitulo = "c. Por t??tulo de Ph. D. o Doctorado equivalente se asignan hasta ochenta (80) puntos. Cuando el docente "
                    + "acredite un t??tulo de Doctorado, y no tenga ning??n t??tulo acreditado de Maestr??a, se le otorgan hasta "
                    + "ciento veinte (120) puntos. No se conceden puntos por t??tulos de Magister o Maestr??a posteriores al "
                    + "reconocimiento de ese doctorado.";
        } else if (titulo.equals("Magister o Maestr??a")) {
            requisitosxTitulo = "b. Por el t??tulo de Magister o Maestr??a se asignan hasta cuarenta (40) puntos.";
        } else if (titulo.equals("Especializaci??n")) {
            requisitosxTitulo = "a. Por t??tulos de Especializaci??n cuya duraci??n est?? entre uno (1) y dos (2) a??os acad??micos, hasta veinte "
                    + "(20) puntos. Por a??o adicional se adjudican hasta diez (10) puntos hasta completar un m??ximo de "
                    + "treinta (30) puntos. Cuando el docente acredite dos (2) especializaciones se computa el n??mero de "
                    + "a??os acad??micos y se aplica lo se??alado en este literal. No se reconocen m??s de dos (2) "
                    + "especializaciones.";
        } else if (titulo.equals("Especializaci??n Cl??nica")) {
            requisitosxTitulo = "PAR??GRAFO II. Para el caso de las especializaciones cl??nicas en medicina humana y odontolog??a, se "
                    + "adjudican quince (15) puntos por cada a??o, hasta un m??ximo acumulable de setenta y cinco (75) puntos.";
        } else if (titulo.equals("Pregrado")) {
            requisitosxTitulo = "a. Por t??tulo de pregrado, ciento setenta y ocho (178) puntos.\n"
                    + "b. Por t??tulo de pregrado en medicina humana o composici??n musical, ciento ochenta y tres (183) puntos.";
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setIndentationLeft(13);
        p.setIndentationRight(12);
        p.setFont(Fonts.SetFont(Color.black, 8, Fonts.ITALIC));
        p.add("\"" + requisitosxTitulo + "\"\n");
        documento.add(p);
        
        if(listaDatosDocentes.get("TIPO_RESOLUCION").equals("Cargo_acad_admin")){
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que por otra parte, el Art??culo 17 del referido decreto, establece los criterios a tener en cuenta para la modificaci??n de salario "+
                    "de los docentes que realizan actividades acad??mico-administrativas, disponi??ndose, en igual sentido, en el Art??culo 62 de la mencionada "+
                    "disposici??n, que el Grupo de Seguimiento al r??gimen salarial y prestacional de los profesores universitarios, puede definir las "+
                    "directrices y criterios que garanticen la homogeneidad, universalidad y coherencia de la informaci??n a nivel nacional, y adem??s adecuar los criterios "+
                    "y efectuar los ajustes a las metodolog??as de evaluaci??n aplicadas por los Comit??s Internos de Asignaci??n de Puntaje o los organismos que hagan sus veces.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que en consulta realizada al Grupo de Seguimiento al R??gimen Salarial y Prestacional de los Profesores Universitarios, bajo Radicado N?? 2011ER62358 "+
                    "sobre el reconocimiento de puntos por productividad acad??mica para los docentes que asuman cargos acad??mico-administrativos, este "+
                    "se??al?? que independientemente de haber elegido entre la remuneraci??n del cargo que va a desempe??ar y la que le corresponde como docente, "+
                    "solo se podr?? hacer efectivo a partir del momento en que el docente termine el ejercicio de las actividades acad??mico-administrativas, de conformidad con el Art??culo 17 del Decreto 1279 de 2002 y el par??grafo 1 del Art??culo 60 del Acuerdo N?? 03 de 2007.\n");
            documento.add(p);
            
            p = new Paragraph(10);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.setAlignment(justificado);
            p.add("Que mediante Resoluci??n Rectoral N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("RESOL_ENCARGO"))+" "+
                    (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+
                    " "+listaDatosDocentes.get("TIPO_VINCULACION")+
                    " "+Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE"))+", "+
                    "fue comisionado para ejercer un cargo de libre nombramiento y remoci??n dentro de la Instituci??n como "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("CARGO"))+" "+
                    "de la Universidad del Magdalena, cargo del que tom?? posesi??n mediante Acta N?? "+
                    Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("ACTA_POSESION"))+".\n");
            documento.add(p);
            
            
        }

        if (titulo.equals("PhD. o Doctorado") && !Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO_MAEST_PUNTAJE")).equals("N/A")) {
            p = new Paragraph(10);
            p.setAlignment(justificado);
            p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
            p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " docente ");
            c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
            p.add(c);
            p.add("se le asignaron cuarenta (40) puntos salariales por el t??tulo de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("TITULO_MAEST_PUNTAJE"))+ ".\n");
            documento.add(p);
        }

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la") + " "
                + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("realiz?? solicitud de asignaci??n de puntos salariales por titulaci??n, aportando copia del diploma de " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_SOLICITUD")) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        try{
        p.add("Que el Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje, en sesi??n realizada el d??a " + fechaEnletras(listaDatosDocentes.get("FECHA_ACTA"),0)
                + " contenida en el numeral "+ listaDatosDocentes.get("NUMERAL_ACTA_CIARP") +" del Acta N?? " + listaDatosDocentes.get("ACTA") + ", asign?? ");
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
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ", ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("por el t??tulo de " + Utilidades.Utilidades.decodificarElemento(tituloCorto[0]) + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("Que mediante acto administrativo motivado expedido por el Rector de UNIMAGDALENA se determina dos (2) veces al a??o el total de puntos que corresponde a cada"
                + " docente, conforme lo se??ala el Art??culo 55 del Decreto 1279 de 2002.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add("En m??rito de lo anterior,\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("RESUELVE:\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO PRIMERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Reconocer y pagar " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")), Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add((listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? ", identificado con " : ", identificada con "));
        if (listaDatosDocentes.get("TIPO_DOC").equals("CC")) {
            p.add("c??dula de ciudadan??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + " expedida en " + Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("C_EXPEDICION")) + ", ");
        } else if (listaDatosDocentes.get("TIPO_DOC").equals("CE")) {
            p.add("c??dula de extranjer??a N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        } else {
            p.add("pasaporte N?? " + FormatoCedula(listaDatosDocentes.get("CEDULA")) + ", ");
        }
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
        p.add(" por el t??tulo de " + Utilidades.Utilidades.decodificarElemento(tituloCorto[0]).replace("\"", "") + ".\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO SEGUNDO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Los ");
        try{
        c = new Chunk(getNumeroDecimal(listaDatosDocentes.get("PUNTOS")) + " (" + ValidarNumeroDec(listaDatosDocentes.get("PUNTOS")) + ") puntos salariales", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
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
            p.add(" surtir??n efectos fiscales una vez "+(listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "el" : "la")+" docente finalice la comisi??n para desempe??ar el cargo de libre nombramiento y remoci??n que le ha sido otorgado.\n");
        }else{
            if(listaDatosDocentes.get("RETROACTIVIDAD").length()==10){
                try{
            p.add(" surtir??n efectos fiscales a partir del d??a " + fechaEnletras(listaDatosDocentes.get("RETROACTIVIDAD"),0) + ", seg??n consta en Acta N?? " + listaDatosDocentes.get("ACTA") + ".\n");
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
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO TERCERO: ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add("Notificar el contenido de la presente decisi??n " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "al" : "a la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Para tal efecto env??esele la correspondiente comunicaci??n al correo electr??nico de notificaci??n, haci??ndole saber que contra la misma procede"
                + " recurso de reposici??n ante el mismo funcionario que la expide, el cual deber?? presentar y sustentar por escrito dentro de los diez (10) d??as"
                + " siguientes a la notificaci??n, de acuerdo con lo preceptuado en los Art??culos 76 y 77 de la Ley 1437 de 2011, "
                + " el Art??culo 55 del Decreto 1279 de 2002 y el Art??culo 47 del Acuerdo Superior N?? 022 de 2019.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO CUARTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" Comunicar la presente decisi??n a la Direcci??n de Talento Humano y al Grupo de N??mina y Prestaciones Sociales a fin de que procedan a realizar"
                + " el tr??mite correspondiente. Env??ese copia de esta resoluci??n al Comit?? Interno de Asignaci??n y Reconocimiento de Puntaje y a la hoja de vida");
        p.add(" " + (listaDatosDocentes.get("SEXO").equalsIgnoreCase("M") ? "del" : "de la") + " " + listaDatosDocentes.get("TIPO_VINCULACION") + " ");
        c = new Chunk(Utilidades.Utilidades.decodificarElemento(listaDatosDocentes.get("NOMBRE_DEL_DOCENTE")) + ".\n ", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        c = new Chunk("ART??CULO QUINTO:", Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(c);
        p.add(" La presente resoluci??n rige a partir del t??rmino de su ejecutoria.\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add("NOTIF??QUESE, COMUN??QUESE Y C??MPLASE\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(" Dada en la ciudad de Santa Marta, D. T. C. H., a los \n\n\n");
        documento.add(p);

        p = new Paragraph(10);
        p.setAlignment(centrado);
        p.setFont(Fonts.SetFont(Color.black, 10, Fonts.BOLD));
        p.add(listaProyecto.get("RECTOR").toUpperCase() + "\n");
        c = new Chunk("Rector\n\n", Fonts.SetFont(Color.black, 10, Fonts.NORMAL));
        p.add(c);
        documento.add(p);

        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Proyect??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("PROYECTO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
        p.add(c);
        p.add(listaProyecto.get("REVISO") + "____");
        documento.add(p);
        p = new Paragraph(8);
        p.setAlignment(justificado);
        p.setFont(Fonts.SetFont(Color.black, 7, Fonts.NORMAL));
        c = new Chunk("Revis??: ", Fonts.SetFont(Color.black, 7, Fonts.BOLD));
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
         * Nombre de los n??meros
         * ************************************************
         */
        Unidades = new String[]{"", "Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Diecis??is", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veinti??n", "Veintid??s", "Veintitr??s", "Veinticuatro", "Veinticinco", "Veintis??is", "Veintisiete", "Veintiocho", "Veintinueve"};
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

            case "Rese??as_Cr??ticas":
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
         * Nombre de los n??meros
         * ************************************************
         */
        Unidades = new String[]{"", "Primero", "Segundo", "Tercero", "Cuarto", "Quinto", "Sexto", "S??ptimo", "Octavo", "Noveno", "D??cimo", "Und??cimo", "Duod??cimo"};
        Decenas = new String[]{"","Decimo", "Vig??simo", "Trig??simo", "Cuadrag??simo", "Quincuag??simo", "Sexag??simo", "Septuag??simo", "Octog??simo", "Nonag??simo"};
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
