/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package Excel;

import Utilidades.Expresiones;
import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.logging.Level; 
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author DOLFHANDLER
 */
public class archivoExcel {
    private File archivoXLS;
    private Workbook libro;
    private String path;

    public static final boolean SI = true, NO = false;

    public archivoExcel(String url) {
        path = url;
        archivoXLS = new File(url);
        libro = new HSSFWorkbook();
    }
    
    public boolean crearArchivo() {
        try {
            if (archivoXLS.exists()) {
                int opcion = JOptionPane.showConfirmDialog(null, "El archivo " + path + " ya se encuentra registrado\n¿Desea reemplazar el archivo existente?");
                if (opcion == JOptionPane.YES_NO_OPTION) {
                    archivoXLS.delete();
                    archivoXLS.createNewFile();
                } else {
                    return false;
                }
            } else {
                archivoXLS.createNewFile();
            }
            return true;
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Ocurrio un problema al tratar de crear el archivo");
            return false;
        }
    }
    
    /**
     * Genera un archivo excel
     * @param contenidoHojas corresponde a los datos que se van a volcar en las hojas del libro excel
     * @param nombreHojas corresponde a los nombres de las hojas del libro
     * @param encabezados corresponde a los encabezados de las celdas en cada hoja
     * @param mostrarOpcionDeAbrir establece si desea ver el documento de inmediato o no
     */
    public void generarArchivoEXCEL(ArrayList<ArrayList<String[]>> contenidoHojas, String[] nombreHojas, ArrayList<String[]> encabezados, boolean mostrarOpcionDeAbrir) {
        try {
            if (crearArchivo()) {
                FileOutputStream archivo = null;

                for (int i = 0; i < contenidoHojas.size(); i++) {
                    archivo = new FileOutputStream(archivoXLS);
                    Sheet hoja = libro.createSheet(nombreHojas[i]);
                    CellStyle estiloDatos = libro.createCellStyle();
                    CellStyle estiloEnc = libro.createCellStyle();

                    Row fila = hoja.createRow(0);
                    Cell celda = null;

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL COLOR DE FONDO DE LAS CELDAS">
                    estiloEnc.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    estiloEnc.setFillForegroundColor(HSSFColor.YELLOW.index);
                    estiloEnc.setAlignment(CellStyle.ALIGN_CENTER);
                    //</editor-fold>

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL TAMAÑO DE LA FUENTE">
                    HSSFFont font = (HSSFFont) libro.createFont();
                    font.setFontName("Calibri");
                    font.setFontHeightInPoints((short) 12);
                    font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                    font.setColor(HSSFColor.BLACK.index);
                    estiloEnc.setFont(font);
                    //</editor-fold>

                    for (int j = 0; j < encabezados.get(i).length; j++) {
                        celda = fila.createCell(j);
                        celda.setCellStyle(estiloEnc);
                        celda.setCellValue(encabezados.get(i)[j]);
                    }

                    //<editor-fold desc="SE COLOCAN LOS BORDES DEL ENCABEZADO">//
                    estiloEnc.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setBottomBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setLeftBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderRight(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setRightBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setTopBorderColor(HSSFColor.BLACK.index);
                    //</editor-fold>

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL TAMAÑO DE LA FUENTE">
                    font = (HSSFFont) libro.createFont();
                    font.setFontName("Calibri");
                    font.setFontHeightInPoints((short) 11);
                    font.setColor(HSSFColor.BLACK.index);
                    estiloDatos.setFont(font);
                    //</editor-fold>

                    for (int j = 0; j < contenidoHojas.get(i).size(); j++) //filas en excel
                    {
                        fila = hoja.createRow(j + 1);
                        for (int k = 0; k < contenidoHojas.get(i).get(j).length; k++) //columnas en excel
                        {
                            celda = fila.createCell(k);
                            celda.setCellValue("" + contenidoHojas.get(i).get(j)[k]);

                            //<editor-fold desc="SE COLOCAN LOS BORDES DE LAS CELDAS">
                            estiloDatos.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setBottomBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setLeftBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderRight(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setRightBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderTop(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setTopBorderColor(HSSFColor.BLACK.index);
                            //</editor-fold>

                            celda.setCellStyle(estiloDatos);
                        }
                    }

                    //SE REALIZA UN AUTO AJUSTE DE LAS CELDAS EN LA HOJA ESPECIFICADA
                    for (int j = 0; j < contenidoHojas.get(i).get(0).length; j++) {
                        hoja.autoSizeColumn(j);
                    }

                }

                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ARCHIVO EXCEL Y SE DA LA OPCION DE ABRIRLO">
                libro.write(archivo);
                archivo.close();
                if (mostrarOpcionDeAbrir) {
                    int opcion = JOptionPane.showConfirmDialog(null, "¿Desea ver el documento " + archivoXLS.getName() + "?\n");
                    if (opcion == JOptionPane.YES_NO_OPTION) {
                        Desktop.getDesktop().open(archivoXLS);
                    }
                }
                //</editor-fold>

            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(archivoExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(archivoExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void generarArchivoEXCELNEW(ArrayList<ArrayList<Map<String, String>>> contenidoHojas, String[] nombreHojas, ArrayList<String[]> encabezados, boolean mostrarOpcionDeAbrir) {
        try {
            if (crearArchivo()) {
                FileOutputStream archivo = null;
                System.out.println("contenido"+contenidoHojas.size());
                for (int i = 0; i < contenidoHojas.size(); i++) {
                    archivo = new FileOutputStream(archivoXLS);
                    Sheet hoja = libro.createSheet(nombreHojas[i]);
                    CellStyle estiloDatos = libro.createCellStyle();
                    CellStyle estiloEnc = libro.createCellStyle();

                    Row fila = hoja.createRow(0);
                    Cell celda = null;

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL COLOR DE FONDO DE LAS CELDAS">
                    estiloEnc.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    estiloEnc.setFillForegroundColor(HSSFColor.YELLOW.index);
                    estiloEnc.setAlignment(CellStyle.ALIGN_CENTER);
                    //</editor-fold>

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL TAMAÑO DE LA FUENTE">
                    HSSFFont font = (HSSFFont) libro.createFont();
                    font.setFontName("Calibri");
                    font.setFontHeightInPoints((short) 12);
                    font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                    font.setColor(HSSFColor.BLACK.index);
                    estiloEnc.setFont(font);
                    //</editor-fold>

                    for (int j = 0; j < encabezados.get(i).length; j++) {
                        celda = fila.createCell(j);
                        celda.setCellStyle(estiloEnc);
                        celda.setCellValue(encabezados.get(i)[j]);
                    }

                    //<editor-fold desc="SE COLOCAN LOS BORDES DEL ENCABEZADO">//
                    estiloEnc.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setBottomBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setLeftBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderRight(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setRightBorderColor(HSSFColor.BLACK.index);
                    estiloEnc.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    estiloEnc.setTopBorderColor(HSSFColor.BLACK.index);
                    //</editor-fold>

                    //<editor-fold defaultstate="collapsed" desc="SE ESTABLECE EL TAMAÑO DE LA FUENTE">
                    font = (HSSFFont) libro.createFont();
                    font.setFontName("Calibri");
                    font.setFontHeightInPoints((short) 11);
                    font.setColor(HSSFColor.BLACK.index);
                    estiloDatos.setFont(font);
                    //</editor-fold>

                    for (int j = 0; j < contenidoHojas.get(i).size(); j++) //filas en excel
                    {
                        fila = hoja.createRow(j + 1);
                        int k = 0;
                        for (Map.Entry<String, String> entry : contenidoHojas.get(i).get(j).entrySet()) {
                            String valor = entry.getValue();
                            celda = fila.createCell(k);
                            celda.setCellValue("" + valor);

                            //<editor-fold desc="SE COLOCAN LOS BORDES DE LAS CELDAS">
                            estiloDatos.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setBottomBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setLeftBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderRight(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setRightBorderColor(HSSFColor.BLACK.index);
                            estiloDatos.setBorderTop(HSSFCellStyle.BORDER_THIN);
                            estiloDatos.setTopBorderColor(HSSFColor.BLACK.index);
                            //</editor-fold>

                            celda.setCellStyle(estiloDatos);
                            k++;
                        }
                    }

                    //SE REALIZA UN AUTO AJUSTE DE LAS CELDAS EN LA HOJA ESPECIFICADA
                    int j = 0;
                    for (Map.Entry<String, String> entry : contenidoHojas.get(i).get(0).entrySet()) {
                        hoja.autoSizeColumn(j);
                        j++;
                    }

                }

                //<editor-fold defaultstate="collapsed" desc="SE CREA EL ARCHIVO EXCEL Y SE DA LA OPCION DE ABRIRLO">
                libro.write(archivo);
                archivo.close();
                if (mostrarOpcionDeAbrir) {
                    int opcion = JOptionPane.showConfirmDialog(null, "¿Desea ver el documento " + archivoXLS.getName() + "?\n");
                    if (opcion == JOptionPane.YES_NO_OPTION) {
                        Desktop.getDesktop().open(archivoXLS);
                    }
                }
                //</editor-fold>

            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(archivoExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(archivoExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
