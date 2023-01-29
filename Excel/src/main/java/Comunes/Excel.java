/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package Comunes;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author USER
 */
public class Excel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
     
            // fuente: 10. Crear, Leer y Modificar Excel en Java y MySQL
            // video https://www.youtube.com/watch?v=oG18JTIKtLo
            
            // Repositori maven: codigo para pom.xml
            // https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/5.2.3
   try {
            crearExcelXLS();
            crearExcelXLSX();
        } catch (IOException ex) {
          //  Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
        public static void crearExcelXLS() throws FileNotFoundException, IOException{
  //  Workbook book=new HSSFWorkbook();
    Workbook book = new HSSFWorkbook();
   Sheet sheet =book.createSheet("Hola Java");
   
   // agregar contenido
   agregar(sheet);
   
   FileOutputStream fileout=new FileOutputStream("TestExcel.xls");
   book.write(fileout);
   fileout.close();
        
        
    }    
    
        
          public static void crearExcelXLSX() throws FileNotFoundException, IOException{
  //  Workbook book=new HSSFWorkbook();
    Workbook book = new HSSFWorkbook();
   Sheet sheet =book.createSheet("Hola Java");

   // agregar contenido
   agregar(sheet);
   
   FileOutputStream fileout=new FileOutputStream("TestExcel.xlsx");
   book.write(fileout);
   fileout.close();
        
        
    }    
        
    public static void agregar(Sheet sheet){
        // ejemplo de agregar: texto, numeros decimales y formula entre celdas   
        Row row=sheet.createRow(0);
            row.createCell(0).setCellValue("Hola Mundo"); // celda A1
            row.createCell(1).setCellValue(7.5); // B1
            row.createCell(2).setCellValue(8.25); // C1
            
            Cell celda = row.createCell(3);
            celda.setCellFormula(String.format("B%d+C%d",1,1)); // sumar B1 y C2
            
            Cell celdaF = row.createCell(4);
            celdaF.setCellFormula(String.format("SUM(B%d:C%d)",1,1)); // sumar B1 y C2
            }    
}
