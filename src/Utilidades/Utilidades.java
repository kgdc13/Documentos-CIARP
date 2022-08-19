package Utilidades;


import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Toolkit;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.imageio.ImageIO;
import javax.swing.DefaultListModel;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;

public class Utilidades {

    public static String TEXTO_SIN_NUMEROS = "\\d";
    public static String SeparadorBusqueda = "__";
    public static String SOLO_NUMEROS = "^[0-9]+$";
    public static String PACIENTES_AUXILIARES_OCASIONALES = "^[A|O]$";
    public static Double FACTOR_CONVERSION = 2.20462;
    public static String DATOS_HISTORICO = "Implante|Descornado|Destete|Hierro fisico|Capado";
    public static String DATOS_PALPACION = "vacia|preñada|repaso";

//    public static void main(String[] args) {
//        Pattern p = Pattern.compile(DATOS_HISTORICO);
//        Matcher m = p.matcher("Hierro fisco");
//        System.out.println("" + m.find());
//    }
    public static boolean validarSINO(String texto) {
        Pattern p = Pattern.compile(DATOS_HISTORICO);
        Matcher m = p.matcher(texto);
        return m.find();
    }

    public static void EstablecerIcono(JFrame vent) {
        vent.setIconImage(Toolkit.getDefaultToolkit().getImage(vent.getClass().getResource("/img/ganaderia32.png")));
    }

    public static void EstablecerPermisosVista(JFrame vista, int idVista) {
        Component[] componentes = vista.getComponents();
        int num = 0;
        for (Component object : componentes) {
            num++;
            System.out.println("******COMPONENTE*****" + num);
        }

    }

    

    public static void guardarImagen(BufferedImage bi, String imagen) {
        try {
            ImageIO.write(bi, Expresiones.obtenerExtension(imagen), new File(imagen));
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "clase: Imagenes\n"
                    + "metodo: ImageIO.write()\n"
                    + "detalle: ocurrio un error al tratar de crear la imagen redimensionada\n"
                    + "" + ex.getMessage());
        }
    }

    public static ArrayList<String> decodificarNombre(String nombre) {
        do {
            nombre = nombre.replace("  ", " ");
        } while (nombre.contains("  "));

        ArrayList<String> resultado = new ArrayList<>();
        String vn[] = nombre.split(" ");

        for (int i = 0; i < vn.length; i++) {
            if (vn[i].equalsIgnoreCase("de")) {
                if (vn[i + 1].equalsIgnoreCase("la") || vn[i + 1].equalsIgnoreCase("los")) {
                    resultado.add(vn[i] + " " + vn[i + 1] + " " + vn[i + 2]);
                    i += 2;
                } else {
                    resultado.add(vn[i] + " " + vn[i + 1]);
                    i += 1;
                }
            } else if (vn[i].equalsIgnoreCase("del")) {
                resultado.add(vn[i] + " " + vn[i + 1]);
                i += 1;
            } else {
                resultado.add(vn[i]);
            }
        }

        vn = null;
        return resultado;
    }

    public static String decodificarElemento(String dato) {
        try {

            dato = "" + dato + "";

            dato = dato.replace("_INTEa_", "¿");
            dato = dato.replace("_INTEc_", "?");

            dato = dato.replace("_at_", "á");
            dato = dato.replace("_et_", "é");
            dato = dato.replace("_it_", "í");
            dato = dato.replace("_ot_", "ó");
            dato = dato.replace("_ut_", "ú");
            dato = dato.replace("_At_", "Á");
            dato = dato.replace("_Et_", "É");
            dato = dato.replace("_It_", "Í");
            dato = dato.replace("_Ot_", "Ó");
            dato = dato.replace("_Ut_", "Ú");

            dato = dato.replace("Ã¡", "á");
            dato = dato.replace("Ã©", "é");
            dato = dato.replace("Ã­", "í");
            dato = dato.replace("Ã³", "ó");
            dato = dato.replace("Ãº", "ú");
            dato = dato.replace("Ã±", "ñ");
            dato = dato.replace("Ã?", "Ñ");

            dato = dato.replace("_enie_", "ñ");
            dato = dato.replace("_ENIE_", "Ñ");

            dato = dato.replace("&ntilde;", "ñ");
            dato = dato.replace("&Ntilde;", "Ñ");

            dato = dato.replace("Â", "");//espacios en blanco
            dato = dato.replace("_CD_", "\"");//COMILLAS DOBLES.."&quot;");//COMILLAS DOBLES..
            dato = dato.replace("_dx_", "<");//menorque..
            dato = dato.replace("_bx_", ">");//menorque..
            dato = dato.replace("_PT_", ":");//2 PUNTOS..
            dato = dato.replace("_M_", "+");//COMILLAS DOBLES..
            dato = dato.replace("_I_", "=");//IGUAL......
            dato = dato.replace("_BS_", "/");// BASESLAS
            dato = dato.replace("_CS_", "\"");//COMILLA SIMPLE...
            dato = dato.replace("_P_", "%");//PORCENTA.....
            //HTML = HTML.replace("_L_n", "<br/>");//ESLAS.....
            dato = dato.replace("_L_", "\\");//ESLAS.....
            dato = dato.replace("_A_", "&");//AMPERSAN.....Â
            dato = dato.replace("_Ord_", "°");//°    
            dato = dato.replace("_ComSpl_", "'");//'    
            //System.out.println("html-->"+HTML);
//            dato = dato.replace("\\n", "<br/>");//° 
            //System.out.println("html2-->"+HTML);
            //HTML = HTML.replace("<br>", "\n");//°
            //HTML = HTML.replace("<br/>", "\n");//°
//            dato = dato.replace("<br>", "\n");//°
//            dato = dato.replace("<br/>", "\n");//° 
            //  alert(HTML);

        } catch (Exception e) {
            System.out.println("ERROR decodificarElemento()-->" + e.toString());
        }

        return dato;
    }

    public static String CodificarElemento(String dato) {
        try {

            dato = "" + dato + "";
            dato = dato.replace("Â", "");//espacios en blanco
            dato = dato.replace("Ã¡", "á");
            dato = dato.replace("Ã©", "é");
            dato = dato.replace("Ã­", "í");
            dato = dato.replace("Ã³", "ó");
            dato = dato.replace("Ãº", "ú");
            dato = dato.replace("Ã±", "ñ");
            dato = dato.replace("Ã?", "Ñ");

            dato = dato.replace("á", "_at_");
            dato = dato.replace("é", "_et_");
            dato = dato.replace("í", "_it_");
            dato = dato.replace("ó", "_ot_");
            dato = dato.replace("ú", "_ut_");

            dato = dato.replace("Á", "_At_");
            dato = dato.replace("É", "_Et_");
            dato = dato.replace("Í", "_It_");
            dato = dato.replace("Ó", "_Ot_");
            dato = dato.replace("Ú", "_Ut_");

            dato = dato.replace("¿", "_INTEa_");
            dato = dato.replace("?", "_INTEc_");

            dato = dato.replace("ñ", "_enie_");
            dato = dato.replace("Ñ", "_ENIE_");

            dato = dato.replace("\"", "_CD_");//COMILLAS DOBLES.."&quot;");//COMILLAS DOBLES..
            dato = dato.replace("<", "_dx_");//menorque..
            dato = dato.replace(">", "_bx_");//menorque..
            dato = dato.replace(":", "_PT_");//2 PUNTOS..
            dato = dato.replace("+", "_M_");//COMILLAS DOBLES..
            dato = dato.replace("=", "_I_");//IGUAL......
            dato = dato.replace("/", "_BS_");// BASESLAS
            dato = dato.replace("\"", "_CS_");//COMILLA SIMPLE...
            dato = dato.replace("%", "_P_");//PORCENTA.....
            //HTML = HTML.replace("_L_n", "<br/>");//ESLAS.....
            dato = dato.replace("\\", "_L_");//ESLAS.....
            dato = dato.replace("&", "_A_");//AMPERSAN.....Â
            dato = dato.replace("°", "_Ord_");//°   
            dato = dato.replace("'", "_ComSpl_");//°   

            //  alert(HTML);
        } catch (Exception e) {
            System.out.println("ERROR decodificarElemento()-->" + e.toString());
        }

        return dato;
    }

    public static boolean validarSoloNumeros(String texto) {
        Pattern p = Pattern.compile(SOLO_NUMEROS);
        Matcher m = p.matcher(texto);
        return m.find();
    }

    public static void validarNumeroEncampodeTexto(JTextField campo) {
        if (validarSoloNumeros(campo.getText())) {
            campo.setText(campo.getText());
        } else {
            campo.setText(campo.getText().substring(0, campo.getText().length() - 1));
        }
    }

    public static void LlenarComboBox(JComboBox combo, List<Map<String, String>> ListaDatos, String KeyMostrar) {
        combo.removeAllItems();
        for (Map<String, String> lista : ListaDatos) {
            if (!lista.containsKey(KeyMostrar)) {
                //return null;
            }
            combo.addItem(Utilidades.decodificarElemento(lista.get(KeyMostrar)));
        }

        //return combo;
    }

    public static void LlenarComboBoxSelected(JComboBox combo, List<Map<String, String>> ListaDatos, String KeyMostrar, String KeySelect, String DatoSelect) {
        combo.removeAllItems();
        for (Map<String, String> lista : ListaDatos) {
            combo.addItem(lista.get(KeyMostrar));
            if (lista.containsKey(KeySelect)) {
                System.out.println("");
                if (lista.get(KeySelect).equals(DatoSelect)) {
                    combo.setSelectedItem(lista.get(KeyMostrar));
                }
            }
        }
        //return combo;
    }

    public static void agregarFilaTabla(DefaultTableModel modelotbl, Object[] fila) {
        modelotbl.addRow(fila);
    }

    public static String MascaraMoneda(String dato) {
        String ret = "";
        int con = 0;
        for (int i = dato.length() - 1; i >= 0; i--) {

            if (con % 3 == 0 && con > 0) {
                ret = "." + ret;
                con = 0;
            }
            ret = "" + dato.charAt(i) + ret;
            con++;

        }
        return ret;
    }

    public static String MascaraMonedaConDecimales(String dato) {
        String ret = "";
        int con = 0;
        String ente = dato.split(",")[0];
        String dec = "", ban = "0";
        if (dato.split(",").length > 1) {
            dec = dato.split(",")[1];
        }
        if (dato.indexOf(",") > -1) {
            ban = "1";
        }
        for (int i = ente.length() - 1; i >= 0; i--) {

            if (con % 3 == 0 && con > 0) {
                ret = "." + ret;
                con = 0;
            }
            ret = "" + ente.charAt(i) + ret;
            con++;

        }
        if (!dec.equals("")) {
            ret = ret + "," + dec;
            ban = "0";
        }
        if (ban.equals("1")) {
            ret = ret + ",";
        }

        return ret;
    }
    public static String MascaraMonedaConDecimalesNeg(String dato) {
        String ret = "";
        int con = 0;
        String ente = dato.split(",")[0];
        String dec = "", ban = "0";
        if (dato.split(",").length > 1) {
            dec = dato.split(",")[1];
        }
        if (dato.indexOf(",") > -1) {
            ban = "1";
        }
        for (int i = ente.length() - 1; i >= 0; i--) {

            if (con % 3 == 0 && con > 0) {
                ret = ((i==0 && (""+ente.charAt(i)).equals("-"))?"":".") + ret;
                con = 0;
            }
            ret = "" + ente.charAt(i) + ret;
            con++;

        }
        if (!dec.equals("")) {
            ret = ret + "," + dec;
            ban = "0";
        }
        if (ban.equals("1")) {
            ret = ret + ",";
        }

        return ret;
    }

    /**
     * Este metodo formatea los numeros escritos en un JTextField.<br>
     * Por ejemplo si el texto escrito es 1234<br>
     * al formatearlo qeedaria asi: 1.234<br>
     * Si el numero escrito fuese 12345,6789<br>
     * al formatearlo qeedaria asi: 12.345,6789<br>
     *
     * @param campoTexto Corresponde al compo de texto sobre el que se dara
     * formato al texto.
     */
    

    public static boolean ValidarEstado(String estado) {
        Pattern p = Pattern.compile(DATOS_PALPACION);
        Matcher m = p.matcher(estado);
        return m.find();
    }

    public String convertirNumeroEnLetras(int numero) {
        String moneda = "";

        if (numero == 1) {
            moneda = " Peso";
        } else {
            moneda = " Pesos";
        }

        return numeroEnLetras(numero) + moneda;

    }

    public static String formatomoneda(String costo) {
        String caux = "";
        String cfin = "";
        int control = 0;
        boolean entero = validarSoloNumeros(costo);

        if (!costo.equals("") && entero) {
            Pattern pat = Pattern.compile("^[0-9]+$");
            Matcher mat = pat.matcher(costo);
            if (mat.matches()) {
                System.out.println("SI");
                for (int i = costo.length() - 1; i >= 0; i--) {
                    control++;
                    caux += costo.charAt(i);
                    if (control % 3 == 0 && i != 0) {
                        caux += ".";
                    }
                }
                for (int j = caux.length() - 1; j >= 0; j--) {
                    cfin += caux.charAt(j);
                }
                return ("$ " + cfin);
            } else {
                return "";
            }
        }
        return "";
    }

    public static String numeroEnLetras(int numero) {
        String[] Unidades, Decenas, Centenas;
        String Resultado = "";

        /**
         * ************************************************
         * Nombre de los números
         * ************************************************
         */
        Unidades = new String[]{"", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veintiuno", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve"};
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
        return Resultado;
    }

    public static String[] obtenerDocumentoyTipoDoc(String TID) {
        String tipoydoc[] = new String[]{"", ""};
        for (int i = 0; i < TID.length(); i++) {
            if (Expresiones.validarSoloNumeros("" + TID.charAt(i))) {
                tipoydoc[0] = TID.substring(0, i);
                tipoydoc[1] = TID.substring(i);
                break;
            }
        }
        return tipoydoc;
    }

    public static String CapitaliceTexto(String texto) {
        try {
            String[] info = texto.trim().toLowerCase().replace(" ", ":").split(":");
            String ret = "";
            String inf = "";
            for (int i = 0; i < info.length; i++) {
                inf = info[i];
                if (!inf.trim().equals("")) {
                    String ini = "" + inf.charAt(0);
                    String fin = inf.substring(1);
                    System.out.println("ini-->" + ini);
                    System.out.println("fin-->" + fin);
                    ret += (ret.equals("") ? "" : " ") + ini.toUpperCase() + fin;
                }
            }
            System.out.println("ret: " + ret);
            return ret;
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    public static List<Map<String, String>> data_list(int caso, List<Map<String, String>> lista, String[] datos) {
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

    public static List<Map<String, String>> data_list(int caso, List<Map<String, String>> lista, String[] datos, String[] datos2) {
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

    public static String CampoNULL(String campo) {
        String ret = "";
        if (campo == null) {
            ret = "NULL";
        } else if (campo.equals("NULL")) {
            ret = "NULL";
        } else {
            ret = "'" + campo + "'";
        }

        return ret;
    }

    public static void LimpiarTabla(JTable tabla) {
        int tam = tabla.getRowCount();
        DefaultTableModel model = (DefaultTableModel) tabla.getModel();
        for (int i = 0; i < tam; i++) {
            model.removeRow(0);
        }
        tabla.setModel(model);
    }

    public static String CapitalizeTexto(String texto) {
        String ret = "";

        if (texto.length() > 0) {
            ret = texto.substring(0, 1).toUpperCase() + texto.substring(1, texto.length());
        }

        return ret;
    }

   

    /**
     * Establece el color de fondo cuando el mouse se posa sobre el jpanel
     * especificado
     *
     * @param panel
     * @param estaEncimaDelPanel si el valor es true, indica que el mouse se
     * encuentra por encima del JPanel
     */
    public static void establecerColorDeFondo(JPanel panel, boolean estaEncimaDelPanel) {
        if (panel.isEnabled()) {
            if (estaEncimaDelPanel) {
                panel.setBackground(new Color(84, 140, 76));
            } else {
                panel.setBackground(new Color(59, 123, 50));
            }
        }
    }

    public static Map<String, String> SplitNames(String name) {
        Map<String, String> retorno = new HashMap<>();
        ArrayList<String> dato = Utilidades.decodificarNombre(name);
        String primero = "", segundo = "";

        if (dato.size() >= 2) {
            primero = dato.get(0);
            segundo = dato.get(1);
            if (dato.size() > 2) {
                for (int i = 2; i < dato.size(); i++) {
                    segundo += " " + dato.get(i);
                }
            }
        } else {
            primero = dato.get(0);
            segundo = "";
        }
        retorno.put("PRIMERO", primero);
        retorno.put("SEGUNDO", segundo);

        return retorno;
    }

    
    public static String ValorNULL(String valor) {
        if (valor == null) {
            valor = "NULL";
        } else if (valor.trim().equals("") || valor.trim().equals("_")) {
            valor = "NULL";
        } else {
            if (!valor.equalsIgnoreCase("now()")) {
                valor = "'" + valor.trim() + "'";
            }
        }
        return valor;
    }

    public static int CompararFechas(String fechaDesde, String fechaHasta) {
        Date fd = new Date(Integer.parseInt(fechaDesde.split("/")[2]) - 1900, Integer.parseInt(fechaDesde.split("/")[1]) - 1, Integer.parseInt(fechaDesde.split("/")[0])),
                fh = new Date(Integer.parseInt(fechaHasta.split("/")[2]) - 1900, Integer.parseInt(fechaHasta.split("/")[1]) - 1, Integer.parseInt(fechaHasta.split("/")[0]));
        int ret = fd.compareTo(fh);
        return ret;
    }

    public static String getIN(String[] datos) {
        String in = "";
        for (String dat : datos) {
            in += (in.equals("") ? "" : ",") + "'" + dat + "'";
        }
        return in;
    }

    public static void LlenarJList(List<Map<String, String>> datos, String keyB, JList Lista, DefaultListModel modlist) {
        LimpiarJList(Lista, modlist);
        for (Map<String, String> dat : datos) {
            modlist.addElement(dat.get(keyB));
        }
        Lista.setModel(modlist);
    }

    public static void LlenarJListByArrayList(ArrayList<String> datos, JList Lista, DefaultListModel modlist) {
        LimpiarJList(Lista, modlist);
        for (String dato : datos) {
            modlist.addElement(dato);
        }
        Lista.setModel(modlist);
    }

    public static String[] getDatosLista(JList lista, List<Map<String, String>> listaDatos, String key) {
        String[] ids = new String[lista.getSelectedIndices().length];
        int x = 0;
        if (lista.isSelectedIndex(0)) {
            ids = new String[lista.getModel().getSize()-1];
            for (int i = 1; i < lista.getModel().getSize(); i++) {
                ids[x++] = listaDatos.get(i).get("" + key);
            }
        } else {
            for (int ind : lista.getSelectedIndices()) {
                ids[x++] = listaDatos.get(ind).get("" + key);
            }
        }
        return ids;
    }

    public static ArrayList<String> getDatosListaByArrayList(JList lista, ArrayList<String> listaDatos) {
        ArrayList<String> salida = new ArrayList<>();
        for (int ind : lista.getSelectedIndices()) {
            salida.add(listaDatos.get(ind));
        }
        return salida;
    }

    public static void LimpiarJList(JList Lista, DefaultListModel modlist) {
        modlist.removeAllElements();
        Lista.setModel(modlist);
    }

   

   
}
