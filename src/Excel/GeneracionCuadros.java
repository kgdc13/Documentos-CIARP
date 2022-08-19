/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package Excel;

import RTF.GeneracionCartas;
import Utilidades.Utilidades;
import com.lowagie.text.DocumentException;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.GroupLayout;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MERRY
 */
public class GeneracionCuadros {
    public String rutaDocentes;
    public String rutaPuntos;
    public String anio;
    public List<Map<String, String>> listaInfoDocentes = new ArrayList<>();
    public List<Map<String, String>> listaInfoPuntos = new ArrayList<>();
    public Map<String, String> InfoParametros = new HashMap<>();
    public Utilidades util = new Utilidades();
    DecimalFormat formateador = new DecimalFormat("#.#");
    
    public GeneracionCuadros(){
        
    }

    public Map<String, String> GenerarCuadros(String anio, String rutaDocentes, String rutaPuntos) throws DocumentException, IOException{
        this.anio = anio;
        this.rutaDocentes = rutaDocentes;
        this.rutaPuntos = rutaPuntos;
        Map<String, String> respuesta = new HashMap<>();
        ControlArchivoExcel con = new ControlArchivoExcel();
        
        //<editor-fold defaultstate="collapsed" desc="Lectura Docentes Planta">
        String extD = rutaDocentes.substring(rutaDocentes.lastIndexOf(".") + 1);
        if (extD.equals("xlsx")) {
            listaInfoDocentes = con.LeerExcelAct(rutaDocentes);
        } else {
            listaInfoDocentes = con.LeerExcel(rutaDocentes);
        }
        if(listaInfoDocentes.size()>0){
            imprimirKeys(listaInfoDocentes.get(0));
        }else{
            respuesta.put("ESTADO", "ERROR");
            respuesta.put("MENSAJE", "Hubo un error con la lectura del archivo de docentes de planta");
            return respuesta;
        }
        
//</editor-fold>
        System.out.println("listaInfoDocentes-->"+listaInfoDocentes.size());
        //<editor-fold defaultstate="collapsed" desc="Lectura Puntos Todos">
        String extP = rutaPuntos.substring(rutaPuntos.lastIndexOf(".") + 1);
        System.out.println("************************FUERA LECTURA DE DOCENTES\n EMPIEZA LECTURA de PUNTOS TODOS");
        System.out.println("rutaDocentes--->"+rutaDocentes);
        System.out.println("rutaPuntos--->"+rutaPuntos);
        System.out.println("extP--->"+extP);
        if (extP.equals("xlsx")) {
            System.out.println("me estoy durmiendo");
            InfoParametros = con.LeerExcelParametrosAct(rutaPuntos, 1, 1,"PUNTOS TODOS");
            System.out.println("AQUI?");
            listaInfoPuntos = con.LeerExcelDesdeAct(rutaPuntos, 2, "PUNTOS TODOS");
            System.out.println("POR POR POR POR POR POR ");
        } else {
            listaInfoPuntos = con.LeerExcelDesde(rutaPuntos, 2, "PUNTOS TODOS");
            System.out.println(" POR ESTE LADO ARRIBA");
            InfoParametros = con.LeerExcelParametros(rutaPuntos, 1, 1); 
            System.out.println("O QUIZA ABAJO");
        }
        
        
//</editor-fold>
        String ruta = "C:\\CIARP\\CUADROS\\";
        String carpetaCiarp = "C:\\CIARP\\";
        String nombreArchivo = "";
        String ext = "xlsx";
        String ext1 = "xls";
        String documento = "";
        System.out.println("listaInfoPuntos--->"+listaInfoPuntos.size());
        System.out.println("listaInfoDocentes-->"+listaInfoDocentes.size());
        
        formatoCedula();
        
        for(Map<String, String> docente : listaInfoDocentes){
            System.out.println("************************************************************************************************");
            System.out.println("*************************DOCENTE*******************"+docente.get("CEDULA")+"**************");
//            System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" + listaInfoDocentes.get(0).get("CEDULA"));
//            System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% 444444" + listaInfoDocentes.get(0).get("NOMBRE_DEL_DOCENTE"));
            List<Map<String, String>> listaDatosDocente = new ArrayList<>();
            listaDatosDocente = data_list(3, listaInfoPuntos, new String[]{"CEDULA<->" + docente.get("CEDULA")});
            System.out.println("-----------------------------listaDatosDocente.size()-->"+listaDatosDocente.size()+"------------------------------------------------------");
            
            nombreArchivo = Capitalize(util.decodificarElemento(docente.get("NOMBRE_DEL_DOCENTE")));
           
            System.out.println("nombre-Archivo--->"+nombreArchivo);
            
            boolean existe = ValidarArchivo(ruta,nombreArchivo, ext1);
            if(existe){try {
                //MODIFICAR ARCHIVO TIPO xls
                
                AgregarCuadroAnioXls(docente, listaDatosDocente, ruta, nombreArchivo, ext1);
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCuadros.class.getName()).log(Level.SEVERE, null, ex);
                    respuesta.put("ESTADO", "ERROR");
                    respuesta.put("MENSAJE", ""+ex.getMessage());
                    respuesta.put("LINEA_ERROR_DOCENTE", ""+docente.get("NOMBRE_DEL_DOCENTE"));
                    return respuesta;
                }
                
            }else if(ValidarArchivo(ruta,nombreArchivo, ext)){try {
                //MODIFICAR ARCHIVO TIPO xlsx
                AgregarCuadroAnioXlsx(docente, listaDatosDocente, ruta, nombreArchivo, ext);
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCuadros.class.getName()).log(Level.SEVERE, null, ex);
                     Logger.getLogger(GeneracionCuadros.class.getName()).log(Level.SEVERE, null, ex);
                    respuesta.put("ESTADO", "ERROR");
                    respuesta.put("MENSAJE", "Error de coincidencia con cuadro anterior "+ex.getMessage());
                    respuesta.put("LINEA_ERROR_DOCENTE", ""+docente.get("NOMBRE_DEL_DOCENTE"));
                    return respuesta;
                }
            }else{try {
                // CREAR ARCHIVO xlsx
                CrearArchivoCruadroXlsx(docente, listaDatosDocente, ruta, nombreArchivo, ext);
                } catch (Exception ex) {
                    Logger.getLogger(GeneracionCuadros.class.getName()).log(Level.SEVERE, null, ex);
                     Logger.getLogger(GeneracionCuadros.class.getName()).log(Level.SEVERE, null, ex);
                    respuesta.put("ESTADO", "ERROR");
                    respuesta.put("MENSAJE", ""+ex.getMessage());
                    respuesta.put("LINEA_ERROR_DOCENTE", ""+docente.get("NOMBRE_DEL_DOCENTE"));
                    return respuesta;
                }
            
            }
            System.out.println("-----------------------------listaDatosDocente.size()-->"+listaDatosDocente.size()+"------------------------------------------------------");
            for(int i=0;i <listaDatosDocente.size();i++){
            System.out.println("lista daos puntos" + listaDatosDocente.get(i).get("PUNTOS"));
            }
        }
        
        System.out.println("**************+FIN DE CREACION DE HOJA PARA LOS DOCENTES");
        JOptionPane.showMessageDialog(null, "Finalizó la creación de cuadros");
    return respuesta;
    }
    
    public void prueba6() throws FileNotFoundException, IOException{
        System.out.println("*******************INICIO PRUEBA***************************");
        String ruta = "C:\\CIARP\\CUADROS\\";
        String carpetaCiarp = "C:\\CIARP\\";
        String docente = "Acosta Salazar Diana Patriciaa";
        String ext = "xlsx";
        String documento = docente+"."+ext;
        
        int anio_ant = Integer.parseInt(anio)-1;
        String nameHoja = "EVAL. " + anio_ant ;
        double puntajeAntetior = 301;
        double puntajeanio = 51;
        ValidarArchivo(ruta, docente, ext);
        System.out.println("ANTES DE ARCHIVO");
        FileInputStream Archivo = new FileInputStream(ruta+documento);
        System.out.println("DESPUES DE ARCHIVO");
        FileInputStream escudo = new FileInputStream(carpetaCiarp+"escudo.png");
        
        System.out.println("ARCHIVO-tam-->"+Archivo.available());
        
        XSSFWorkbook libro = new XSSFWorkbook(Archivo);
        
        XSSFSheet hojaAnterior = libro.getSheet(nameHoja);
        
        if(libro.getSheetIndex("EVAL. "+anio)>-1){
            libro.removeSheetAt(libro.getSheetIndex("EVAL. "+anio));
        }
        
        XSSFSheet hojaNueva = libro.createSheet("EVAL. "+anio);
        ArrayList<String> ListaEstProf = getDatosHojaAnterior(hojaAnterior, "1. ESTUDIOS PROFESIONALES", "2. ESTUDIOS DE POSTGRADO");
        ArrayList<String> ListaEstPost = getDatosHojaAnterior(hojaAnterior, "2. ESTUDIOS DE POSTGRADO", "3. CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:");
        
        
        Row filaAnterior = hojaAnterior.getRow(12);
        Cell celdaAnterior  = filaAnterior.getCell(11);
        
        puntajeAntetior = celdaAnterior.getNumericCellValue();
        
        Font fontBold = libro.createFont();
        fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        Font fontNormal = libro.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        
        ///////////////////////////////HOJA NUEVA
        CellStyle style = libro.createCellStyle();
        style.setFont(fontBold);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        
        int numfila = 8;
        XSSFRow newrow = hojaNueva.createRow(numfila);
        XSSFCell newcell = newrow.createCell(0);
        newcell.setCellValue("");
        
        addImageN(hojaNueva,escudo);
        
        
        CrearCelda("", newrow, 1, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("COMITÉ INTERNO DE ASIGNACIÓN Y RECONOCIMIENTO DE PUNTAJE", newrow, 2, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 12, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(8, 8, 2, 11));   
        
        //LimpiarBordes(style);
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        EstiloBorde(5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, style);
        CrearCelda("", newrow, 1, 5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("APLICACIÓN DEL DECRETO 1279 DE 2002", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(9, 9, 2, 11)); 
        CrearCelda("", newrow, 12, 6, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 11, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 2, 7)); 
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 9, 11)); 
        
        //<editor-fold defaultstate="collapsed" desc="Datos Docente">
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("DOCENTE:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("_NOMBRE-DOCENTE_", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Fecha de evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("_FECHA-EVAL_", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(11, 16, 7, 8)); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("IDENTIFICACIÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("identificacion", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje actual:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+puntajeAntetior, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("valor punto", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("INGRESO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("ingreso", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje esta evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("punt anual", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("facul xxxxx", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("ANTIGÜEDAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("antigüedad", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("punt total", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
//        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("valor punto", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Salario 1279:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("pesopunto*puntos", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("facul xxxxx", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
//        CrearCelda("Salario total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("lo mismo", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        //</editor-fold>
        
        //SEPARADOR
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(17, 17, 2, 11)); 
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        //<editor-fold defaultstate="collapsed" desc="Titulacion Docente">
        int item = 1;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ESTUDIOS PROFESIONALES", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Fecha Reconocimiento", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("Puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(String datosProf: ListaEstProf){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda(""+datosProf, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ESTUDIOS DE POSTGRADO", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(String datosPost: ListaEstPost){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda(""+datosPost, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". CATEGORIA ACADEMICA EN EL ESCALAFON:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda("Categoria-", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". EXPERIENCIA CALIFICADA:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda("Evaluac. Desempeño año "+anio, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("2", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". PRODUCCION ACADEMICA SALARIAL:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);    

        for(int i = 0; i < 5;i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("Produccion academica N°"+i, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda("fecha", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ACTIVIDADES DE DIRECCION:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("Actividad n° 1", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("fecha", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("TOTAL", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("tot_punt", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". BONIFICACION POR PRODUCTIVIDAD", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Puntos", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        
        for(int i = 0; i < 5;i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("Bonificacion N°"+i, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda("Puntos B", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        }
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldasBlanco(newrow, 1, 12, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(int i = 0; i < 2; i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldasBlanco(newrow, 1, 12, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        //<editor-fold defaultstate="collapsed" desc="FIRMA VICERECTOR">
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("VoBo. "+"NOMBRE_VIERRECTOR", newrow, 4, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 5, 6, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("Vicerrector Académico", newrow, 4, 0, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        //Vicerrector Académico
        //</editor-fold>
        
        
        
        
        //</editor-fold>
        
        
        
        EstablecerTamanioColumnasHoja(hojaNueva);
        
        Archivo.close();
        
         FileOutputStream salida = new FileOutputStream(ruta+documento);
         libro.write(salida);
         salida.close();
        
        
        System.out.println("******************END METHOD***************************");
        
        
    }
    
    /**
     * 
     * @param Borde
     * 0 -> none 
     * 1 -> top
     * 2 -> bottom
     * 3 -> left
     * 4 -> right
     * 5 -> top, left
     * 6 -> top, right
     * 7 -> top, bottom
     * 8 -> bottom, left
     * 9 -> bottom right
     * 10-> left, right
     * 11-> top, left, right
     * 12-> bottom left, right
     * 13-> top, left, bottom
     * 14-> top, right, bottom
     * 15-> tpo, left, bottom, right
     * @param tipoBorde
     * short thin, double, thick
     * @param estilo 
     * cellstyle para adicionar estilo
     */
    public void EstiloBorde(int Borde, short tipoBorde, short color, CellStyle estilo){
        
        if(Borde == 1 || Borde == 5 || Borde == 6 || Borde == 7 || Borde == 11 || Borde == 13 || Borde == 14 || Borde == 15){
            estilo.setBorderTop(tipoBorde);
            estilo.setTopBorderColor(color);
        }
        
        if(Borde == 2 || Borde == 7 || Borde == 8 || Borde == 9 || Borde == 12 || Borde == 13 || Borde == 14 || Borde == 15){
            estilo.setBorderBottom(tipoBorde);
            estilo.setBottomBorderColor(color);
        }
        
        if(Borde == 3 || Borde == 5 || Borde == 8 || Borde == 10 || Borde == 11 || Borde == 12 || Borde == 13 || Borde == 15){
            estilo.setBorderLeft(tipoBorde);
            estilo.setLeftBorderColor(color);
        }
        
        if(Borde == 4 || Borde == 6 || Borde == 9 || Borde == 10 || Borde == 11 || Borde == 12 || Borde == 14 || Borde == 15){
            estilo.setBorderRight(tipoBorde);
            estilo.setRightBorderColor(color);
        }
    }
    
    public void EstiloRegionBorde(int Borde, short tipoBorde, CellRangeAddress rango, short color, XSSFSheet hoja, XSSFWorkbook libro){
        
        if(Borde == 1 || Borde == 5 || Borde == 6 || Borde == 7 || Borde == 11 || Borde == 13 || Borde == 14 || Borde == 15){
            RegionUtil.setBorderTop(tipoBorde, rango, hoja, libro);
            RegionUtil.setTopBorderColor(color, rango, hoja, libro);
        }
        
        if(Borde == 2 || Borde == 7 || Borde == 8 || Borde == 9 || Borde == 12 || Borde == 13 || Borde == 14 || Borde == 15){
            RegionUtil.setBorderBottom(tipoBorde, rango, hoja, libro);
            RegionUtil.setBottomBorderColor(color, rango, hoja, libro);
        }
        
        if(Borde == 3 || Borde == 5 || Borde == 8 || Borde == 10 || Borde == 11 || Borde == 12 || Borde == 13 || Borde == 15){
            RegionUtil.setBorderLeft(tipoBorde, rango, hoja, libro);
            RegionUtil.setLeftBorderColor(color, rango, hoja, libro);
        }
        
        if(Borde == 4 || Borde == 6 || Borde == 9 || Borde == 10 || Borde == 11 || Borde == 12 || Borde == 14 || Borde == 15){
            RegionUtil.setBorderRight(tipoBorde, rango, hoja, libro);
            RegionUtil.setRightBorderColor(color, rango, hoja, libro);
        }
    }
    
    public void CrearCeldaOLD(String Descripcion, XSSFRow fila, int col, CellStyle style){
        XSSFCell celda = fila.createCell(col);
        celda.setCellValue(Descripcion);
        celda.setCellStyle(style);
    }
    
    /**
     * 
     * @param Descripcion
     * @param fila
     * @param col
     * @param borde
     * @param tipoBorde
     * @param color
     * @param aling
     * @param font
     * @param libro 
     */
    public void CrearCelda(String Descripcion, XSSFRow fila, int col, int borde, short tipoBorde, short color, short aling, Font font, XSSFWorkbook libro){
        CellStyle style = libro.createCellStyle();
        style.setFont(font);
        style.setAlignment(aling);
//        style.setWrapText(true);
        
        EstiloBorde(borde, tipoBorde, color, style);
        XSSFCell celda = fila.createCell(col);
        celda.setCellValue(Descripcion);
        celda.setCellStyle(style);
    }
    
    public void CrearCeldasBlanco(XSSFRow fila, int colInicial, int numCols, int borde, short tipoBorde, short color, short aling, Font font, XSSFWorkbook libro){
        CellStyle style = libro.createCellStyle();
        style.setFont(font);
        style.setAlignment(aling);
        style.setWrapText(true);
        EstiloBorde(borde, tipoBorde, color, style);
        for(int i = colInicial; i <= numCols; i++){
            XSSFCell newcell = fila.createCell(i);
            newcell.setCellValue("");
            newcell.setCellStyle(style);
        }
    }
    public void CrearCeldasBlanco(XSSFRow fila, int colInicial, int numCols, CellStyle style){
        for(int i = colInicial; i <= numCols; i++){
            XSSFCell newcell = fila.createCell(i);
            newcell.setCellValue("");
            newcell.setCellStyle(style);
        }
    }
    
    public void EstiloCelda(){
        
    }

    private void EstablecerTamanioColumnasHoja(XSSFSheet hoja) {
        hoja.setColumnWidth(0, 767);
        hoja.setColumnWidth(1, 730);
        hoja.setColumnWidth(2, 4712);
        hoja.setColumnWidth(3, 2009);
        hoja.setColumnWidth(4, 4712);
        hoja.setColumnWidth(5, 1644);
        hoja.setColumnWidth(6, 1534);
        hoja.setColumnWidth(7, 2922);
        hoja.setColumnWidth(8, 1278);
        hoja.setColumnWidth(9, 3543);
        hoja.setColumnWidth(10, 1826);
        hoja.setColumnWidth(11, 2885);
        hoja.setColumnWidth(12, 840);
    }

    private void LimpiarBordes(CellStyle estilo) {
        estilo.setBorderTop(CellStyle.BORDER_NONE);
    }

    private CellStyle RenovarEstilo(Font font, XSSFWorkbook libro) {
        CellStyle style = libro.createCellStyle();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(font);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        return style;
    }

    private ArrayList<String> getDatosHojaAnterior(XSSFSheet hojaAnterior, String Dato1, String Dato2) {
        //System.out.println("****************************getDatosHojaAnterior***********************************+");
        ArrayList<String> retorno =  new ArrayList<>();
        int[] filas = getFilasEntreItem(hojaAnterior, Dato1, Dato2);
        for(int i = filas[0]; i <= filas[1]; i++){
            Row fila = hojaAnterior.getRow(i);
            Cell celda  = fila.getCell(3);
            String dato = celda.getStringCellValue();
            if(!dato.equals(""))
                retorno.add(dato);
        }
        return retorno;
    }

    private ArrayList<String> getDatosHojaAnteriorOld(Sheet hojaAnterior, String Dato1, String Dato2) {
        //System.out.println("****************************getDatosHojaAnterior***********************************+");
        ArrayList<String> retorno =  new ArrayList<>();
        int[] filas = getFilasEntreItemOld(hojaAnterior, Dato1, Dato2);
        for(int i = filas[0]; i <= filas[1]; i++){
            Row fila = hojaAnterior.getRow(i);
            Cell celda  = fila.getCell(3);
            String dato = celda.getStringCellValue();
            retorno.add(dato);
        }
        return retorno;
    }
    
    private int[] getFilasEntreItem(XSSFSheet hojaAnterior, String Dato1, String Dato2) {
//        System.out.println("***********getFilasEntreItem****************"+hojaAnterior);
        System.out.println("Dato1-RSRD-<"+Dato1+">--");
        System.out.println("Dato2-RSRD-<"+Dato2+">--");
        int[] filas = new int[2];
        int filaRef = 15;
        boolean encontroPrimera = false;
        boolean encontroSegunda = false;
        int cel = 2;
        while(!encontroPrimera){
            Row fila = hojaAnterior.getRow(filaRef);
            Cell celda  = fila.getCell(cel);
            
            if(celda != null){
                System.out.println("celda.getStringCellValue()--<"+celda.getStringCellValue());
                if(celda.getStringCellValue().equals(Dato1)){
                    encontroPrimera = true;
                    filas[0] = filaRef+1;
                }
            }
            filaRef++;
        }
//        System.out.println("******************------------------------------------********************");
        while(!encontroSegunda){
            Row fila = hojaAnterior.getRow(filaRef);
            Cell celda  = fila.getCell(2);
//            System.out.println("celda.getStringCellValue()--<"+celda.getStringCellValue());
            if(celda.getStringCellValue().equals(Dato2)){
                encontroSegunda = true;
                filas[1] = filaRef-1;
            }
            filaRef++;
        }
//        System.out.println("fila[0]-->"+filas[0]);
//        System.out.println("fila[1]-->"+filas[1]);
//        
//        System.out.println("***********END getFilasEntreItem****************");
        
        return filas;
    }

    private int[] getFilasEntreItemOld(Sheet hojaAnterior, String Dato1, String Dato2) {
        //System.out.println("***********getFilasEntreItem****************");
        int[] filas = new int[2];
        int filaRef = 15;
        boolean encontroPrimera = false;
        boolean encontroSegunda = false;
        while(!encontroPrimera){
            Row fila = hojaAnterior.getRow(filaRef);
            Cell celda  = fila.getCell(2);
            if(celda.getStringCellValue().equals(Dato1)){
                encontroPrimera = true;
                filas[0] = filaRef+1;
            }
            filaRef++;
        }
        
        while(!encontroSegunda){
            Row fila = hojaAnterior.getRow(filaRef);
            Cell celda  = fila.getCell(2);
            if(celda.getStringCellValue().equals(Dato2)){
                encontroSegunda = true;
                filas[1] = filaRef-1;
            }
            filaRef++;
        }
//        System.out.println("fila[0]-->"+filas[0]);
//        System.out.println("fila[1]-->"+filas[1]);
//        
//        System.out.println("***********END getFilasEntreItem****************");
        
        return filas;
    }
    
    protected void addImage(Sheet sheet, FileInputStream escudo) throws IOException {
        
        BufferedImage img = ImageIO.read(escudo);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(img, ".png", baos);
        baos.flush();
        byte[] imageInByte = baos.toByteArray();
        baos.close();
        //Revisar si es png o jpg
        int pictureIdx = sheet.getWorkbook().addPicture(imageInByte, XSSFWorkbook.PICTURE_TYPE_PNG);
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(2);
        anchor.setRow1(12);
        anchor.setRow2(17);
        anchor.setCol2(3);

        Picture pict = drawing.createPicture(anchor, pictureIdx);
        
    //pict.resize();
        
         //
        }
    protected void addImageN(Sheet sheet, FileInputStream escudo) throws IOException {
//         read the image to the stream 
        final CreationHelper helper = sheet.getWorkbook().getCreationHelper(); 
        
        final Drawing drawing = sheet.createDrawingPatriarch(); 
        final ClientAnchor anchor = helper.createClientAnchor(); 
        anchor.setAnchorType( ClientAnchor.MOVE_AND_RESIZE ); 
        final int pictureIndex = sheet.getWorkbook().addPicture(IOUtils.toByteArray(escudo), XSSFWorkbook.PICTURE_TYPE_JPEG); 
        anchor.setCol1(5);
        anchor.setRow1(2);
        anchor.setRow2(8);
        anchor.setCol2(7);
        final Picture pict = drawing.createPicture(anchor, pictureIndex); 
        pict.resize(1.2); 
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
                    System.out.println("lista---->"+lista.size());
                    for (Map<String, String> lis : lista) {
                        int coincidencias = 0;
                        for (String prm : datos) {
                            String[] item = prm.split("<->");
//                            System.out.println("prm....>"+prm);
//                            System.out.println("lis.get(item[0])--->"+lis.get(item[0]));
//                            System.out.println("item[1]-->"+item[1]+"//");
                            if (lis.get(item[0]).equals(item[1])) {
//                                System.out.println("if");
                                coincidencias++;
                                
                            }
                        }
//                        System.out.println("coincidencias--->"+coincidencias);
                        if (coincidencias == datos.length) {
                            rlista.add(lis);
                        }
                    }
                    break;
                }
                case 30: {
                    for (String prm : datos) {
                        String[] item = prm.split("<->");
                        for (Map<String, String> lis : lista) {
                            if (lis.get(item[0]).equals(item[1])) {
                                rlista.add(lis);
                            }
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
                    int fil = 0;
                    for (Map<String, String> lis : lista) {
                        fil++;
                        int coincidencias = 0;
                        for (String prm : datos2) {
                            String[] item = prm.split("<->");
                            boolean t = lis.containsKey(item[0]);
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

    private boolean ValidarArchivo(String ruta, String nombre, String ext) throws FileNotFoundException, IOException {
        System.out.println("******ValidarArchivo*****");
        System.out.println("*"+ruta+nombre+ext);
        File file = new File(ruta+nombre+"."+ext);
        if (file.exists()) {// si el archivo existe se elimina
            System.out.println("Archivo Encontrado");
            return true;
        }
        System.out.println(" Archivo NO Encontrado");
        
        return false;
    }

    private String Capitalize(String texto) {
        try {
            String[] info = texto.trim().toLowerCase().replace(" ", ":").split(":");
            String ret = "";
            String inf = "";
            for (int i = 0; i < info.length; i++) {
                inf = info[i];
                if (!inf.trim().equals("")) {
                    String ini = "" + inf.charAt(0);
                    String fin = inf.substring(1);
                    ret += (ret.equals("") ? "" : " ") + ini.toUpperCase() + fin;
                }
            }
            return ret;
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    private void AgregarCuadroAnioXlsx(Map<String, String> docente, List<Map<String, String>> listaDatosDocente, String ruta, String nombreArchivo, String ext) throws FileNotFoundException, IOException, Exception {
        System.out.println("*******************AgregarCuadroAnioXlsx***************************");
        System.out.println("*************************************************"+docente.get("NOMBRE_DEL_DOCENTE")+"**********************************************");
        String carpetaCiarp = "C:\\CIARP\\";
        String documento = nombreArchivo+"."+ext;
        
        List<Map<String, String>> listaBonificaciones = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Bonificacion"});
        List<Map<String, String>> listaTitulacion = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Titulacion"});
        List<Map<String, String>> categoria = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Ascenso_en_el_escalafon"});
        List<Map<String, String>> listaSalarial = data_list(30, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Salarial", "TIPO_RESOLUCION<->Salarial_colciencias"});
        
        List<Map<String, String>> listaTitulacionCarg = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Titulacion"});
        List<Map<String, String>> listaAscensoCarg = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin", "TIPO_PRODUCTO<->Ascenso_en_el_Escalafon_Docente"});
        List<Map<String, String>> listaCargAcad = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Cargo_acad_admin"});
        
        
//       if(listaDatosDocente.size()>0){
//        for(int i=0;i>=listaDatosDocente.size();i++){
//            System.out.println("DATOS DE PUNTOS DE LISTA SALARIALS"+Double.parseDouble(listaDatosDocente.get(i).get("PUNTOS")));
//        }
//        }
        System.out.println("listaSalarial--->"+listaSalarial.size());
        System.out.println("listaTitulacion--->"+listaTitulacion.size());
        System.out.println("listaTitulacionCarg--->"+listaTitulacionCarg.size());
        System.out.println("listaAscensoCarg--->"+listaAscensoCarg.size());
        System.out.println("listaCargAcad--->"+listaCargAcad.size());
        if(listaTitulacionCarg.size()>0){
            listaTitulacion.addAll(listaTitulacionCarg);
        }
        if(listaAscensoCarg.size()>0){
            categoria.addAll(listaAscensoCarg);
        }
        
        String indx = "";
        
        if(listaCargAcad.size()>0){
            for(int i = 0; i < listaCargAcad.size(); i++){
                if(listaCargAcad.get(i).get("TIPO_PRODUCTO").equals("Titulacion")||listaCargAcad.get(i).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")){
                     indx += (indx.equals("")?"":"<::>")+""+i;
                                        
                }
//                else if(listaCargAcad.get(i).get("TIPO_PRODUCTO").equals("Ascenso_en_el_Escalafon_Docente")){
//                    indx += (indx.equals("")?"":"<::>")+""+i;
//                }
            }
            if(!indx.equals("")){
                String[] ind = indx.split("<::>");
                for(int i = ind.length-1; i >= 0; i--){
                    
                    listaCargAcad.remove(Integer.parseInt(ind[i]));
                }
            }
          
            if(listaCargAcad.size()>0){
                listaSalarial.addAll(listaCargAcad);
            }
        } 
      
        System.out.println("listaSalarial--->"+listaSalarial.size());
        System.out.println("listaTitulacion--->"+listaTitulacion.size());
       
        
        
        double puntajeAntetior = 301;
        double puntajeanio = getPuntaje(listaSalarial, listaTitulacion, categoria);
        
        
//        String antiguedad = getAntiguedad(docente.get("DIAVIN")+"-"+docente.get("MESVIN")+"-"+docente.get("AÑOVIN"));
        FileInputStream Archivo = new FileInputStream(ruta+documento);
        FileInputStream escudo = new FileInputStream(carpetaCiarp+"escudo.png");
        System.out.println("RUTA + DOCUMENTO----"+ruta+documento);
        
        XSSFWorkbook libro = new XSSFWorkbook(Archivo);
         
        String nameHoja = "EVAL. 2019";
        System.out.println("libro.getNumberOfSheets()--->"+libro.getNumberOfSheets());
        if(libro.getNumberOfSheets()>1){
            nameHoja = "EVAL. "+(Integer.parseInt(anio)-1);
        }else{
            nameHoja = "EVAL. Inicial";
        }
        System.out.println("name-<"+nameHoja);
        XSSFSheet hojaAnterior = libro.getSheet(nameHoja);
        
        if(libro.getSheetIndex("EVAL. "+anio)>-1){
            libro.removeSheetAt(libro.getSheetIndex("EVAL. "+anio));
        }
        System.out.println("elimando hoja -- EVAL. "+anio );
        
        XSSFSheet hojaNueva = libro.createSheet("EVAL. "+anio);
        String uno = "1. ESTUDIOS PROFESIONALES";
        String dos = "2. ESTUDIOS DE POSTGRADO";
        
        if(docente.get("TIPO_PERFIL").equals("SIN TITULO UNIVERSITARIO")){
            uno = "1. FORMACIÓN ACADÉMICA BÁSICA (CAP. VIII, DECRETO 1279 DE 2002";
            dos = "2. POR ESTUDIOS UNIVERSITARIOS RELACIONADOS";						
				

        }
        
        ArrayList<String> ListaEstProf = getDatosHojaAnterior(hojaAnterior, uno, dos);
        ArrayList<String> ListaEstPost = getDatosHojaAnterior(hojaAnterior, dos, "3. CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:");
        ArrayList<String> EstCategoria = getDatosHojaAnterior(hojaAnterior, "3. CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:", "4. EXPERIENCIA CALIFICADA:");
        
        
        Row filaAnterior = hojaAnterior.getRow(14);
        Cell celdaAnterior  = filaAnterior.getCell(11);
        if(celdaAnterior.getCellType() == celdaAnterior.CELL_TYPE_NUMERIC){
            puntajeAntetior = celdaAnterior.getNumericCellValue();
        }else if(celdaAnterior.getCellType() == celdaAnterior.CELL_TYPE_FORMULA){
            puntajeAntetior = celdaAnterior.getNumericCellValue();
        }else{
            String p = celdaAnterior.getStringCellValue();
            System.out.println("p--->"+p);
            puntajeAntetior = Double.parseDouble(p);
        }
        
        System.out.println("puntajeAntetior--->"+puntajeAntetior);
        System.out.println("puntajeanio--->"+puntajeanio);
        String puntajeTotal = ""+(puntajeAntetior+puntajeanio);
        String ValorPesoPunto = ValidarNumeroDec(InfoParametros.get("PESO_PUNTO"));
        System.out.println("punt--->"+ValorPesoPunto);
        double salario = Double.parseDouble(ValorPesoPunto) * Double.parseDouble(puntajeTotal);
        System.out.println("Salario--->"+salario+"<<<<");
        String salarioF = ValidarNumeroDec(""+salario);
        System.out.println("puntajeTotal--->"+puntajeTotal);
        puntajeTotal = ValidarNumeroDec(puntajeTotal);
        System.out.println("puntajeTotal--->"+puntajeTotal);
        Font fontBold = libro.createFont();
        fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        Font fontNormal = libro.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        
        ///////////////////////////////HOJA NUEVA
        CellStyle style = libro.createCellStyle();
        style.setFont(fontBold);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        
        int numfila = 8;
        XSSFRow newrow = hojaNueva.createRow(numfila);
        XSSFCell newcell = newrow.createCell(0);
        newcell.setCellValue("");
        
        addImageN(hojaNueva,escudo);
        
        
        CrearCelda("", newrow, 1, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("COMITÉ INTERNO DE ASIGNACIÓN Y RECONOCIMIENTO DE PUNTAJE", newrow, 2, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 12, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(8, 8, 2, 11));   
        
        //LimpiarBordes(style);
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        EstiloBorde(5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, style);
        CrearCelda("", newrow, 1, 5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("APLICACIÓN DEL DECRETO 1279 DE 2002", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(9, 9, 2, 11)); 
        CrearCelda("", newrow, 12, 6, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 11, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 2, 7)); 
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 9, 11)); 
        
        //<editor-fold defaultstate="collapsed" desc="Datos Docente">
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("DOCENTE:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+Utilidades.decodificarElemento(docente.get("NOMBRE_DEL_DOCENTE")), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Fecha de evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("31/12/"+anio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(11, 16, 7, 8)); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("IDENTIFICACIÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+docente.get("CEDULA"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje actual:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(ValidarNumeroDec(""+puntajeAntetior), newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("INGRESO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+docente.get("DIAVIN")+"/"+docente.get("MESVIN")+"/"+docente.get("AÑOVIN"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+ValorPesoPunto, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje esta evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(ValidarNumeroDec(""+puntajeanio), newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+Capitalize(docente.get("FACULTAD")), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("ANTIGÜEDAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+antiguedad, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+puntajeTotal, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Salario 1279:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+salarioF, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
//        numfila++;
//        newrow = hojaNueva.createRow(numfila);
//        CrearCelda("cc", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
////        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
////        CrearCelda(""+Capitalize(docente.get("FACULTAD")), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
////////        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
////        CrearCelda("Salario total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
////        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
////        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
////        CrearCelda(""+salarioF, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("dd", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        
        //</editor-fold>
        
        //SEPARADOR
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(16, 16, 2, 11)); 
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        //<editor-fold defaultstate="collapsed" desc="Titulacion Docente">
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER,fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(uno, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Fecha Reconocimiento", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("Puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(String datosProf: ListaEstProf){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda(""+datosProf, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(dos, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        System.out.println("ListaEstPost--->"+ListaEstPost.size());
        for(String datosPost: ListaEstPost){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda(""+datosPost, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        if(listaTitulacion.size()>0){
            System.out.println("Pintar listaTitulacion-->"+listaTitulacion.size());
            for(Map<String, String> titulacion: listaTitulacion){
                numfila++;
                newrow = hojaNueva.createRow(numfila);
                CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
                CrearCelda(""+Utilidades.decodificarElemento(titulacion.get("NOMBRE_SOLICITUD")), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
                CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
                CrearCelda(""+Utilidades.decodificarElemento(titulacion.get("FECHA_ACTA")), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
                CrearCelda(""+titulacion.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
                CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                }
        }
        
        int item = 2;
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
       
        String cate = "";

        if(categoria.size()>0){
           System.out.println("Pintar listaAScenso-->"+categoria.size());
            for(Map<String, String> categ: categoria){
                numfila++;
                newrow = hojaNueva.createRow(numfila);
                CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
                CrearCelda(""+Utilidades.decodificarElemento(categ.get("NOMBRE_SOLICITUD")), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
                 CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
                CrearCelda(""+Utilidades.decodificarElemento(categ.get("FECHA_ACTA")), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
                CrearCelda(""+categ.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
                CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                }
        }else{
            cate = EstCategoria.get(0);
             numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda(""+Utilidades.decodificarElemento(cate), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". EXPERIENCIA CALIFICADA:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda("Evaluac. Desempeño año "+anio, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("2", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". PRODUCCIÓN ACADÉMICA SALARIAL:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);    

        System.out.println("Lista ----> Salarial------"+listaSalarial.size());
        
        for(Map<String, String> salarial: listaSalarial){
            numfila++;
            String descripcion = "";
            if(Utilidades.decodificarElemento(salarial.get("SUBTIPO_PRODUCTO")).equals("N/A")){
                descripcion = salarial.get("TIPO_PRODUCTO")+":";
            }else{
                descripcion = salarial.get("SUBTIPO_PRODUCTO")+":";
            }
            descripcion += " "+ getNOMBREPRODUCTO(salarial); 
            System.out.println("descripcion-->"+descripcion);
            System.out.println("INTENTADO DESCIFRAR REL PORQUE DEL VALOR SIN DECIMAL" +ValidarNumeroDec(salarial.get("PUNTOS").replace(",",".")));
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda(""+Utilidades.decodificarElemento(descripcion).replace("_", " "), newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda(""+Utilidades.decodificarElemento(salarial.get("FECHA_ACTA")), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda(ValidarNumeroDec(salarial.get("PUNTOS").replace(",", ".")), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }   
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ACTIVIDADES DE DIRECCIÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        

        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);

        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("TOTAL", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(""+puntajeanio, newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". BONIFICACIÓN POR PRODUCTIVIDAD", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Puntos", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        
        if(listaBonificaciones.size()>0){
            for(Map<String, String> bonificacion: listaBonificaciones){
                String descripcion = "";
                if(Utilidades.decodificarElemento(bonificacion.get("SUBTIPO_PRODUCTO")).equals("N/A")){
                    descripcion = Utilidades.decodificarElemento(bonificacion.get("TIPO_PRODUCTO"))+":";
                }else{
                    descripcion = Utilidades.decodificarElemento(bonificacion.get("SUBTIPO_PRODUCTO"))+":";
                }
                descripcion += " "+ getNOMBREPRODUCTO(bonificacion); 

                numfila++;
                newrow = hojaNueva.createRow(numfila);

                CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda(""+Utilidades.decodificarElemento(descripcion).replace("_", " "), newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
                CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
                CrearCelda(ValidarNumeroDec(bonificacion.get("PUNTOS")), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
                CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
                CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
            }
        }else{
            numfila++;
            newrow = hojaNueva.createRow(numfila);

            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        }
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldasBlanco(newrow, 1, 12, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(int i = 0; i < 2; i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldasBlanco(newrow, 1, 12, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        //<editor-fold defaultstate="collapsed" desc="FIRMA VICERECTOR">
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("VoBo. "+InfoParametros.get("VICERRECTOR"), newrow, 4, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 5, 6, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda(""+InfoParametros.get("LABEL_VICERRECTOR")+"\n", newrow, 4, 0, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        //Vicerrector Académico
        //</editor-fold>
        
        
        
        
        //</editor-fold>
        
        
        
        EstablecerTamanioColumnasHoja(hojaNueva);
        
        Archivo.close();
        
         FileOutputStream salida = new FileOutputStream(ruta+documento);
         libro.write(salida);
         salida.close();
        
        
        System.out.println("******************END METHOD***************************");
    }

    private void CrearArchivoCruadroXlsx(Map<String, String> docente, List<Map<String, String>> listaDatosDocente, String ruta, String nombreArchivo, String ext) throws FileNotFoundException, IOException, Exception {
        System.out.println("*******************CrearArchivoCruadroXlsx***************"+docente.get("NOMBRE_DEL_DOCENTE")+"************");
        String carpetaCiarp = "C:\\CIARP\\";
        String documento = nombreArchivo+"."+ext;
        
        List<Map<String, String>> listaBonificaciones = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Bonificacion"});
        List<Map<String, String>> listaTitulacion = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Titulacion"});
        List<Map<String, String>> categoria = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Ascenso_en_el_escalafon"});
        List<Map<String, String>> listaSalarial = data_list(30, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Salarial", "TIPO_RESOLUCION<->Salarial_colciencias", "TIPO_RESOLUCION<->Cargo_acad_admin"});
        
        
        
        double puntajeanio = getPuntaje(listaSalarial, listaTitulacion, categoria);
        String ValorPesoPunto = "1234567";
        double salario = Double.parseDouble(ValorPesoPunto) * puntajeanio;
        
//        String antiguedad = getAntiguedad(docente.get("DIAVIN")+"-"+docente.get("MESVIN")+"-"+docente.get("AÑOVIN"));
        
        FileInputStream escudo = new FileInputStream(carpetaCiarp+"escudo.png");
                   
        
        XSSFWorkbook libro = new XSSFWorkbook();
         
        XSSFSheet hojaNueva = libro.createSheet("EVAL. Inicial");
        
        
        Font fontBold = libro.createFont();
        fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        Font fontNormal = libro.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        
        ///////////////////////////////HOJA NUEVA
        CellStyle style = libro.createCellStyle();
        style.setFont(fontBold);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        
        int numfila = 8;
        XSSFRow newrow = hojaNueva.createRow(numfila);
        XSSFCell newcell = newrow.createCell(0);
        newcell.setCellValue("");
        
        addImageN(hojaNueva,escudo);
        
        
        CrearCelda("", newrow, 1, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("COMITÉ INTERNO DE ASIGNACION Y RECONOCIMIENTO DE PUNTAJE", newrow, 2, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 12, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(8, 8, 2, 11));   
        
        //LimpiarBordes(style);
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        EstiloBorde(5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, style);
        CrearCelda("", newrow, 1, 5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("APLICACIÓN DEL DECRETO 1279 DE 2002", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(9, 9, 2, 11)); 
        CrearCelda("", newrow, 12, 6, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 11, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 2, 7)); 
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 9, 11)); 
        
        //<editor-fold defaultstate="collapsed" desc="Datos Docente">
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("DOCENTE:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+docente.get("NOMBRE_DEL_DOCENTE"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Fecha de evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("31/12/"+anio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(11, 16, 7, 8)); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("IDENTIFICACION:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+docente.get("CEDULA"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje actual:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+ValorPesoPunto, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("INGRESO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+docente.get("DIAVIN")+"/"+docente.get("MESVIN")+"/"+docente.get("AÑOVIN"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje esta evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+puntajeanio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda(""+docente.get("FACULTAD"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda("ANTIGÜEDAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+antiguedad, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Puntaje total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+puntajeanio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+ValorPesoPunto, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCelda("Salario 1279:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda(""+salario, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCelda(""+docente.get("FACULTAD"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        CrearCeldasBlanco(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
//        CrearCelda("Salario total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCeldasBlanco(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
//        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
//        CrearCelda(""+salario, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
//        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        //</editor-fold>
        
        //SEPARADOR
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(17, 17, 2, 11)); 
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        //<editor-fold defaultstate="collapsed" desc="Titulacion Docente">
        int item = 1;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ESTUDIOS PROFESIONALES", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Fecha Reconocimiento", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("Puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda("ESTUDIO ROFESIONAL", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ESTUDIOS DE POSTGRADO", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
            
        
        
        if(listaTitulacion.size()>0){
            for(Map<String, String> titulacion: listaTitulacion){
                numfila++;
                newrow = hojaNueva.createRow(numfila);
                CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
                CrearCelda(""+titulacion.get("NOMBRE_SOLICITUD"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
                CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
                CrearCelda(""+titulacion.get("FECHA_ACTA"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda(""+titulacion.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                }
        }else{
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCelda("Estudios Postgrados", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        String cate = "";
        if(categoria.size()>0){
            cate = categoria.get(0).get("NOMBRE_SOLICITUD");
        }else{
            cate = "Profesor Asistente";
        }
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda(""+cate, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(""+cate, newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". EXPERIENCIA CALIFICADA:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCelda("Evaluac. Desempeño año "+anio, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("2", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". PRODUCCIÓN ACADÉMICA SALARIAL:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);    

        for(Map<String, String> salarial: listaSalarial){
            numfila++;
            String descripcion = "";
            if(salarial.get("SUBTIPO_PRODUCTO").equals("N/A")){
                descripcion = salarial.get("TIPO_PRODUCTO")+":";
            }else{
                descripcion = salarial.get("SUBTIPO_PRODUCTO")+":";
            }
            descripcion += " "+ salarial.get("NOMBRE_SOLICITUD"); 
            
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda(""+descripcion, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda(""+salarial.get("FECHA_ACTA"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda(""+salarial.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }   
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". ACTIVIDADES DE DIRECCIÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        

        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);

        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("TOTAL", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("tot_punt", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("", newrow, 9, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda(item+". BONIFICACIÓN POR PRODUCTIVIDAD", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCelda("Puntos", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        
        for(Map<String, String> bonificacion: listaBonificaciones){
            String descripcion = "";
            if(bonificacion.get("SUBTIPO_PRODUCTO").equals("N/A")){
                descripcion = bonificacion.get("TIPO_PRODUCTO")+":";
            }else{
                descripcion = bonificacion.get("SUBTIPO_PRODUCTO")+":";
            }
            descripcion += " "+ bonificacion.get("NOMBRE_SOLICITUD"); 
            
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("B"+descripcion, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlanco(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCelda(""+bonificacion.get("PUNTOS"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCelda("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
            CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        }
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 9, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 10, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCelda("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldasBlanco(newrow, 1, 12, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(int i = 0; i < 2; i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldasBlanco(newrow, 1, 12, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        //<editor-fold defaultstate="collapsed" desc="FIRMA VICERECTOR">
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda("VoBo. "+InfoParametros.get("VICERRECTOR"), newrow, 4, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlanco(newrow, 5, 6, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCelda(""+InfoParametros.get("LABEL_VICERRECTOR"), newrow, 4, 0, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        //Vicerrector Académico
        //</editor-fold>
        
        
        
        
        //</editor-fold>
        
        
        
        EstablecerTamanioColumnasHoja(hojaNueva);
        
        
         FileOutputStream salida = new FileOutputStream(ruta+documento);
         libro.write(salida);
         salida.close();
        
        
        System.out.println("******************END METHOD***************************");
    }

    private void AgregarCuadroAnioXls(Map<String, String> docente, List<Map<String, String>> listaDatosDocente, String ruta, String nombreArchivo, String ext) throws FileNotFoundException, IOException, Exception {
        System.out.println("*******************AgregarCuadroAnioXls***********"+docente.get("NOMBRE_DEL_DOCENTE")+"****************");
        String carpetaCiarp = "C:\\CIARP\\";
        String documento = nombreArchivo+"."+ext;
        
        List<Map<String, String>> listaBonificaciones = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Bonificacion"});
        List<Map<String, String>> listaTitulacion = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Titulacion"});
        List<Map<String, String>> categoria = data_list(3, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Ascenso_en_el_escalafon"});
        List<Map<String, String>> listaSalarial = data_list(30, listaDatosDocente, new String[]{"TIPO_RESOLUCION<->Salarial", "TIPO_RESOLUCION<->Salarial_colciencias", "TIPO_RESOLUCION<->Cargo_acad_admin"});
        
        
        
        double puntajeAntetior = 301;
        double puntajeanio = getPuntaje(listaSalarial, listaTitulacion, categoria);
        String puntajeTotal = ""+(puntajeAntetior+puntajeanio);
        String ValorPesoPunto = "1234567";
        double salario = Double.parseDouble(ValorPesoPunto) * puntajeanio;
        
//        String antiguedad = getAntiguedad(docente.get("DIAVIN")+"-"+docente.get("MESVIN")+"-"+docente.get("AÑOVIN"));
        FileInputStream Archivo = new FileInputStream(ruta+documento);
        FileInputStream escudo = new FileInputStream(carpetaCiarp+"escudo.jpg");
                   
        
        HSSFWorkbook libro = new HSSFWorkbook(Archivo);
         
        String nameHoja = "EVAL. 2019";
        if(libro.getNumberOfSheets()>1){
            nameHoja = "EVAL. "+(Integer.parseInt(anio)-1);
        }else{
            nameHoja = "EVAL. Inicial";
        }
            
        Sheet hojaAnterior = libro.getSheet(nameHoja);
        
        if(libro.getSheetIndex("EVAL. "+anio)>-1){
            libro.removeSheetAt(libro.getSheetIndex("EVAL. "+anio));
        }
         
        Sheet hojaNueva = libro.createSheet("EVAL. "+anio);
        ArrayList<String> ListaEstProf = getDatosHojaAnteriorOld(hojaAnterior, "1. ESTUDIOS PROFESIONALES", "2. ESTUDIOS DE POSTGRADO");
        ArrayList<String> ListaEstPost = getDatosHojaAnteriorOld(hojaAnterior, "2. ESTUDIOS DE POSTGRADO", "3. CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:");
        ArrayList<String> EstCategoria = getDatosHojaAnteriorOld(hojaAnterior, "3. CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:", "4. EXPERIENCIA CALIFICADA:");
        
        
        Row filaAnterior = hojaAnterior.getRow(12);
        Cell celdaAnterior  = filaAnterior.getCell(11);
        
        puntajeAntetior = celdaAnterior.getNumericCellValue();
        
        Font fontBold = libro.createFont();
        fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        Font fontNormal = libro.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        
        ///////////////////////////////HOJA NUEVA
        CellStyle style = libro.createCellStyle();
        style.setFont(fontBold);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        
        int numfila = 8;
        Row newrow = hojaNueva.createRow(numfila);
        Cell newcell = newrow.createCell(0);
        newcell.setCellValue("");
        
        addImageN(hojaNueva,escudo);
        
        
        CrearCeldaOld("", newrow, 1, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("COMITÉ INTERNO DE ASIGNACION Y RECONOCIMIENTO DE PUNTAJE", newrow, 2, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 12, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(8, 8, 2, 11));   
        
        //LimpiarBordes(style);
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        EstiloBorde(5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, style);
        CrearCeldaOld("", newrow, 1, 5, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("APLICACIÓN DEL DECRETO 1279 DE 2002", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(9, 9, 2, 11)); 
        CrearCeldaOld("", newrow, 12, 6, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 2, 11, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 2, 7)); 
        hojaNueva.addMergedRegion(new CellRangeAddress(10, 10, 9, 11)); 
        
        //<editor-fold defaultstate="collapsed" desc="Datos Docente">
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("DOCENTE:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld(""+docente.get("NOMBRE_DEL_DOCENTE"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Fecha de evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("31/12/"+anio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(11, 16, 7, 8)); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("IDENTIFICACION:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld(""+docente.get("CEDULA"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Puntaje actual:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld(""+puntajeAntetior, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("INGRESO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld(""+docente.get("DIAVIN")+"/"+docente.get("MESVIN")+"/"+docente.get("AÑOVIN"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Puntaje esta evaluación:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld(""+puntajeanio, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("ANTIGÜEDAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld("", newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Puntaje total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld(""+puntajeTotal, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("VALOR PESO PUNTO:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld(""+ValorPesoPunto, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Salario 1279:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld(""+salario, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("FACULTAD:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld(""+docente.get("FACULTAD"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 6, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 6)); 
        CrearCeldaOld("Salario total:", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 10, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld(""+salario, newrow, 11, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        //</editor-fold>
        
        //SEPARADOR
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 9, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 2, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        hojaNueva.addMergedRegion(new CellRangeAddress(17, 17, 2, 11)); 
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        //<editor-fold defaultstate="collapsed" desc="Titulacion Docente">
        int item = 1;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". ESTUDIOS PROFESIONALES", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("Fecha Reconocimiento", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("Puntos", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(String datosProf: ListaEstProf){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCeldaOld(""+datosProf, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCeldasBlancoOld(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". ESTUDIOS DE POSTGRADO", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(String datosPost: ListaEstPost){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
            CrearCeldaOld(""+datosPost, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlancoOld(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
            CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        if(listaTitulacion.size()>0){
            for(Map<String, String> titulacion: listaTitulacion){
                numfila++;
                newrow = hojaNueva.createRow(numfila);
                CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
                CrearCeldaOld(""+titulacion.get("NOMBRE_SOLICITUD"), newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
                CrearCeldasBlancoOld(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
                CrearCeldaOld(""+titulacion.get("FECHA_ACTA"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCeldaOld(""+titulacion.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
                }
        }
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". CATEGORÍA ACADÉMICA EN EL ESCALAFÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        String cate = "";
        if(categoria.size()>0){
            cate = categoria.get(0).get("NOMBRE_SOLICITUD");
        }else{
            cate = EstCategoria.get(0);
        }
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldaOld(""+cate, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(""+cate, newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". EXPERIENCIA CALIFICADA:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldaOld("Evaluac. Desempeño año "+anio, newrow, 3, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 4, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 3, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("2", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". PRODUCCIÓN ACADÉMICA SALARIAL:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);    

        for(Map<String, String> salarial: listaSalarial){
            numfila++;
            String descripcion = "";
            if(salarial.get("SUBTIPO_PRODUCTO").equals("N/A")){
                descripcion = salarial.get("TIPO_PRODUCTO")+":";
            }else{
                descripcion = salarial.get("SUBTIPO_PRODUCTO")+":";
            }
            descripcion += " "+ salarial.get("NOMBRE_SOLICITUD"); 
            
            newrow = hojaNueva.createRow(numfila);
            CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld(""+descripcion, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlancoOld(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCeldaOld(""+salarial.get("FECHA_ACTA"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCeldaOld(""+salarial.get("PUNTOS"), newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }   
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". ACTIVIDADES DE DIRECCIÓN:", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        

        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
        CrearCeldasBlancoOld(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);

        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("TOTAL", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("tot_punt", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 2, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("", newrow, 9, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 7, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        //CrearCelda("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        
        item++;
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld(item+". BONIFICACIÓN POR PRODUCTIVIDAD", newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
        CrearCeldaOld("Puntos", newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
        CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        
        for(Map<String, String> bonificacion: listaBonificaciones){
            String descripcion = "";
            if(bonificacion.get("SUBTIPO_PRODUCTO").equals("N/A")){
                descripcion = bonificacion.get("TIPO_PRODUCTO")+":";
            }else{
                descripcion = bonificacion.get("SUBTIPO_PRODUCTO")+":";
            }
            descripcion += " "+ bonificacion.get("NOMBRE_SOLICITUD"); 
            
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 8, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("B"+descripcion, newrow, 2, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_LEFT, fontNormal, libro);
            CrearCeldasBlancoOld(newrow, 3, 7, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 2, 7)); 
            CrearCeldaOld(""+bonificacion.get("PUNTOS"), newrow, 9, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
            CrearCeldaOld("", newrow, 10, 15, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 9, 10)); 
            CrearCeldaOld("", newrow, 11, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
            CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);  
        }
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("", newrow, 1, 3, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 9, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 10, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldaOld("", newrow, 12, 4, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro); 
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldasBlancoOld(newrow, 1, 12, 1, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        
        for(int i = 0; i < 2; i++){
            numfila++;
            newrow = hojaNueva.createRow(numfila);
            CrearCeldasBlancoOld(newrow, 1, 12, 0, CellStyle.BORDER_MEDIUM, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        }
        
        //<editor-fold defaultstate="collapsed" desc="FIRMA VICERECTOR">
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld("VoBo. "+InfoParametros.get("VICERRECTOR"), newrow, 4, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        CrearCeldasBlancoOld(newrow, 5, 6, 1, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontBold, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        
        numfila++;
        newrow = hojaNueva.createRow(numfila);
        CrearCeldaOld(""+InfoParametros.get("LABEL_VICERRECTOR"), newrow, 4, 0, CellStyle.BORDER_THIN, IndexedColors.BLACK.index, CellStyle.ALIGN_CENTER, fontNormal, libro);
        hojaNueva.addMergedRegion(new CellRangeAddress(numfila, numfila, 4, 6));
        //Vicerrector Académico
        //</editor-fold>
        
        
        
        
        //</editor-fold>
        
        
        
        EstablecerTamanioColumnasHojaOld(hojaNueva);
        
        Archivo.close();
        
         FileOutputStream salida = new FileOutputStream(ruta+documento+"x");
         libro.write(salida);
         salida.close();
        
        
        System.out.println("******************END METHOD***************************");
    }

    

    private String getAntiguedady(String fechaVinculacion) {
        String fechaEvaluacion = "31-12-"+anio;
        System.out.println("antiguedad*********************************************");
        System.out.println("*"+fechaVinculacion+"*");
        if(fechaVinculacion.equals("_-_-_")){
            return "";
        }
        String antiguedad = "";
        
        String[] fechVin=fechaVinculacion.split("-");
        String[] fechEval=fechaEvaluacion.split("-");
        
        int meses = Integer.parseInt(fechVin[1]) - Integer.parseInt(fechEval[1]);
        int redanio=0;
        if(meses > 0){
            meses = 12-meses;
            redanio = 1;
        }else{
            meses = meses * -1;
        }
        
        
        int anios = Integer.parseInt(fechEval[2]) - Integer.parseInt(fechVin[2]) - redanio;
        
        if(anios>0){
            if(anios > 1)
                antiguedad = anios +" años";
            else
                antiguedad = anios +" año";
        }
        if(meses > 0){
            if(!antiguedad.equals("")){
                antiguedad +=" y ";
            }
            if(meses > 1)
                antiguedad += meses +" meses";
            else
                antiguedad += meses +" mes";
        }
        
        
        return antiguedad;
    }

    private double getPuntaje(List<Map<String, String>> listaSalarial, List<Map<String, String>> listaTitulacion, List<Map<String, String>> listaCategoria) throws Exception{
        double puntaje = 2;
        
        for(Map<String, String> lista: listaTitulacion){
            puntaje += Double.parseDouble(lista.get("PUNTOS").replace(",", "."));
        }
        for(Map<String, String> lista: listaSalarial){
            puntaje += Double.parseDouble(lista.get("PUNTOS").replace(",", "."));
        }
        for(Map<String, String> lista: listaCategoria){
            puntaje += Double.parseDouble(lista.get("PUNTOS").replace(",", "."));
        }
        
        return puntaje;
    }


    
    public void CrearCeldaOld(String Descripcion, Row fila, int col, int borde, short tipoBorde, short color, short aling, Font font, HSSFWorkbook libro){
        CellStyle style = libro.createCellStyle();
        style.setFont(font);
        style.setAlignment(aling);
        
        EstiloBorde(borde, tipoBorde, color, style);
        Cell celda = fila.createCell(col);
        celda.setCellValue(Descripcion);
        celda.setCellStyle(style);
    }
    
    public void CrearCeldasBlancoOld(Row fila, int colInicial, int numCols, int borde, short tipoBorde, short color, short aling, Font font, HSSFWorkbook libro){
        CellStyle style = libro.createCellStyle();
        style.setFont(font);
        style.setAlignment(aling);
        EstiloBorde(borde, tipoBorde, color, style);
        for(int i = colInicial; i <= numCols; i++){
            Cell newcell = fila.createCell(i);
            newcell.setCellValue("");
            newcell.setCellStyle(style);
        }
    }
    public void CrearCeldasBlancoOld(Row fila, int colInicial, int numCols, CellStyle style){
        for(int i = colInicial; i <= numCols; i++){
            Cell newcell = fila.createCell(i);
            newcell.setCellValue("");
            newcell.setCellStyle(style);
        }
    }
    

    private void EstablecerTamanioColumnasHojaOld(Sheet hoja) {
        hoja.setColumnWidth(0, 767);
        hoja.setColumnWidth(1, 730);
        hoja.setColumnWidth(2, 4712);
        hoja.setColumnWidth(3, 2009);
        hoja.setColumnWidth(4, 4712);
        hoja.setColumnWidth(5, 1644);
        hoja.setColumnWidth(6, 1534);
        hoja.setColumnWidth(7, 2922);
        hoja.setColumnWidth(8, 1278);
        hoja.setColumnWidth(9, 3543);
        hoja.setColumnWidth(10, 1826);
        hoja.setColumnWidth(11, 2885);
        hoja.setColumnWidth(12, 840);
    }


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
        System.out.println("numero----->"+valor);
        if(valor.indexOf(",") > -1){
            
            valor = valor.replace(",", ".");
        }
        
        
        if (valor.indexOf(".") > -1) {
            
            System.out.println("numero------>"+valor);
            Double dat = Double.parseDouble(valor);
            System.out.println("dat------>"+dat);
            DecimalFormat df = new DecimalFormat("0.00");
            System.out.println("df.format(dat)------>"+df.format(dat));
            valor = df.format(dat);
            System.out.println("numero------>"+valor);
            valor = valor.replace(".", ",");
            String[] daot= valor.split(",");
            System.out.println("DAOT [0]" +daot[0]);
//            if (daot[0].equals("")){
//                daot[0]= "0";
//                valor = daot[0]+valor;
//                System.out.println("DAO[0] = "+daot[0]);
//                System.out.println("DAO[1] = "+daot[1]);
//            }
            if(daot[1].equals("00")){
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

    private void formatoCedula() {
        for(int i = 0; i < listaInfoPuntos.size(); i++){
            Map<String, String> dat = listaInfoPuntos.get(i);
            String d = ""+(long)Double.parseDouble(dat.get("CEDULA").replace(",", "."));
//            System.out.println("d-->"+d);
            dat.put("CEDULA", d);
            listaInfoPuntos.set(i, dat);
        }
    }
    
     private String getNOMBREPRODUCTO(Map<String, String> datos) {
        String datosProducto = "";
        switch (datos.get("TIPO_PRODUCTO")) {
//            case "Ingreso_a_la_Carrera_Docente":
//                datosSoporte = ComillasSoporte(listadatosdocentexTipoProducto.get(k).get("SOPORTES"));
//                respuestaSoporte = datosSoporte;
//                break;
//            case "Ascenso_en_el_Escalafon_Docente":
//                datosSoporte = ComillasSoporte(listadatosdocentexTipoProducto.get(k).get("SOPORTES"));
//                respuestaSoporte = datosSoporte;
//                break;
            case "Articulo":
                datosProducto =  " "+Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; de la revista " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
                case "Art_Col":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; de la revista " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Libro":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; editorial " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Capitulo_de_Libro":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; editorial " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISBN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Ponencias_en_Eventos_Especializados":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; en el " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Publicaciones_Impresas_Universitarias":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; de la revista " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + "; " +ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Reseñas_Críticas":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; de la revista " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Traducciones":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                        + "; de la revista " + Utilidades.decodificarElemento(datos.get("REVISTA/EVENTO/EDITORIAL"))
                        + "; ISSN: " + Utilidades.decodificarElemento(datos.get("ISSN/ISBN"))
                        + "; " + Utilidades.decodificarElemento(datos.get("FECHA_PUBLICACION/REALIZACION"))
                        + " (" + Utilidades.decodificarElemento(datos.get("PUBLINDEX"))
                        + "); " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                break;
            case "Direccion_de_Tesis":
                datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD")) + " ";
                break;
            default:
                if (!Utilidades.decodificarElemento(datos.get("N_AUTORES")).equals("N/A")) {
                    datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                            + "; " + ValidarNumeroDec(datos.get("N_AUTORES")) + " autor(es).";
                } else {
                    datosProducto = " " + Utilidades.decodificarElemento(datos.get("NOMBRE_SOLICITUD"))
                            + " ";
                }
                break;
        }

        return datosProducto;
    } 

    private void imprimirKeys(Map<String, String> dato) {
//        System.out.println("***LECTURA DE KEYS DE LISTA*******");
//        for (Map.Entry<String, String> entry : dato.entrySet()) {
//            String key = entry.getKey();
//            
//            System.out.println("key: "+key);
//            
//        }
//        System.out.println("***END LECTURA DE KEYS DE LISTA*******\n");
    }
}
