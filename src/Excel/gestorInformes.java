/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Excel;

import Utilidades.Expresiones;
import com.toedter.calendar.JCalendar;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import javax.swing.JOptionPane;

/**
 *
 * @author DOLFHANDLER
 */
public class gestorInformes extends Thread implements Runnable {

    private Map<String, String> list;
    private ArrayList<ArrayList<Map<String, String>>> datos = new ArrayList<>();
    private Thread proceso;
    private archivoExcel excel;

    public gestorInformes(Map<String, String> list, ArrayList<ArrayList<Map<String, String>>> datos) {
        this.list = list;
        this.datos = datos;
//        Calendar cal = Calendar.getInstance();
//        SimpleDateFormat sdf = new SimpleDateFormat("ddMMyyyy");
        //String sufijo = sdf.format(cal.getTime())+""+((int)((Math.random() * 9999999) + 10000000));
        excel=new archivoExcel(Expresiones.guardarEn()+"\\"+list.get("NOMBRE_ARCHIVO")+list.get("NUMACTA")+".xls");
    }

    public synchronized void iniciar() {
        proceso = new Thread(this, "proceso informes");
        proceso.start();
    }

    public synchronized void terminar() {
        try {   
            proceso.join();
        } catch (InterruptedException ex) {
            JOptionPane.showMessageDialog(null, "terminar -> " + ex.getMessage());
        }
    }

    @Override
    public void run() {
        if (list != null) {
            ArrayList<String[]> encabezados = new ArrayList<>();
            String enc = "";
            System.out.println("LOCALIZANDO ERRORES"+enc);
            System.out.println("datos.size->"+datos.size());
            System.out.println("datos.get(0)->"+datos.get(0).size());
            for (Map.Entry<String, String> entry : datos.get(0).get(0).entrySet()) {
                String key = entry.getKey();
                enc += (enc.equals("")?"":"<:-:>")+key;
            }
            encabezados.add(enc.split("<:-:>"));
            System.out.println("");
            excel.generarArchivoEXCELNEW(datos, new String[]{""+list.get("NOMBRE_ARCHIVO").replace("_","")}, encabezados, true);
        }
    }
}
