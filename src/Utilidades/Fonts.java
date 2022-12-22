/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package Utilidades;

import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import java.awt.Color;
import java.io.IOException;

/**
 *
 * @author KENNY
 */
public class Fonts extends Font{
    
    /**
     *
     */
    public static BaseFont arialFont;

    public Fonts(BaseFont font){
        arialFont = font;
    }
    public static Font arial = new Font(arialFont);
    public static Font arialSiete = new Font(arialFont, Font.NORMAL, 7, Color.BLACK);
    public static Font arialBoldSiete = new Font(arialFont, Font.BOLD, 7, Color.BLACK); 
    
    public static Font arialItalicOcho = new Font(arialFont, Font.ITALIC, 8, Color.BLACK);
    public static Font arialBoldOcho = new Font(arialFont, Font.BOLD, 8, Color.BLACK);
    public static Font arialOcho = new Font(arialFont, Font.NORMAL, 8, Color.BLACK);
    
    public static Font arialDiez = new Font(arialFont, Font.NORMAL, 10, Color.BLACK);
    public static Font arialBoldDiez = new Font(arialFont, Font.BOLD, 10, Color.BLACK); 
    
    public static Font arialOnce = new Font(arialFont, Font.NORMAL, 11, Color.BLACK);
    public static Font arialBoldOnce = new Font(arialFont, Font.BOLD, 11, Color.BLACK);
    public static Font arialUnderLineOnce = new Font(arialFont, Font.UNDERLINE, 11, Color.BLACK);

    public static Font SetFont(Color color, float tamanio, int estilo){
        Font fuente = new Font(arialFont);
        fuente.setColor(color);
        fuente.setSize(tamanio);
        fuente.setStyle(estilo);
        return fuente;
    }
    public static Font SetFontTwoStyle(Color color, float tamanio, int estilo){
        Font fuente = new Font(arialFont);
        fuente.setColor(color);
        fuente.setSize(tamanio);
        fuente.setStyle(estilo);
        fuente.setStyle("underlined");
        return fuente;
    }
}
