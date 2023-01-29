/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package Comunes;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        // Repositorio maven: codigo para pom.xml
        // https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/5.2.3
        try {
            //Crear un EXCEL
            crearExcelXLS();
            crearExcelXLSX();

            //lectura de EXCEL
            leerXLS();
            leerXLSX();
            //Modificar fichero existente en EXCEL
            ModificaXLS();
            ModificaXLSX();
        } catch (IOException ex) {
            //  Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void crearExcelXLS() throws FileNotFoundException, IOException {
        //  Workbook book=new HSSFWorkbook();
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Hola Java");

        // agregar contenido
        agregar(sheet);

        FileOutputStream fileout = new FileOutputStream("TestExcel.xls");
        book.write(fileout);
        fileout.close();

    }

    public static void crearExcelXLSX() throws FileNotFoundException, IOException {

        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Hola Java");

        // agregar contenido
        agregar(sheet);

        FileOutputStream fileout = new FileOutputStream("TestExcel.xlsx");
        book.write(fileout);
        fileout.close();

    }

    public static void agregar(Sheet sheet) {
        // ejemplo de agregar: texto, numeros decimales y formula entre celdas   
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Hola Mundo"); // celda A1
        row.createCell(1).setCellValue(7.5); // B1
        row.createCell(2).setCellValue(8.25); // C1

        Cell celda = row.createCell(3);
        celda.setCellFormula(String.format("B%d+C%d", 1, 1)); // sumar B1 y C2

        Cell celdaF = row.createCell(4);
        celdaF.setCellFormula(String.format("SUM(B%d:C%d)", 1, 1)); // sumar B1 y C2
    }

    public static void leerXLS() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream("LeerExcel.xls");

        HSSFWorkbook wb = new HSSFWorkbook(file);

        HSSFSheet sheet = wb.getSheetAt(0);
        int numFilas = sheet.getLastRowNum();
        System.out.println("Leyendo fichero: xls");
        for (int a = 0; a <= numFilas; a++) {
            Row fila = sheet.getRow(a);
            int numCols = fila.getLastCellNum();
            for (int b = 0; b < numCols; b++) {
                Cell celda = fila.getCell(b);

                switch (celda.getCellType().toString()) {
                    case "NUMERIC":
                        System.out.print(celda.getNumericCellValue() + " ");
                        break;
                    case "STRING":
                        System.out.print(celda.getStringCellValue() + " ");
                        break;
                    case "FORMULA":
                        System.out.print(celda.getCellFormula() + " ");
                        break;
                }
            }
            System.out.println("");
        }
    }

    public static void ModificaXLS() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream("LeerExcel.xls");

        HSSFWorkbook wb = new HSSFWorkbook(file);

        HSSFSheet sheet = wb.getSheetAt(0);

        // comprobar que exista la fila y la celda...
        HSSFRow fila = sheet.getRow(1);

        if (fila == null) {
            fila = sheet.createRow(1);

        }

        HSSFCell celda = fila.getCell(1); //leer la celda para conservar formato.
        if (celda == null) {
            celda = fila.createCell(1);
        }

        celda.setCellValue("Modificacion");

        file.close();

        FileOutputStream output = new FileOutputStream("Modificado.xls");
        wb.write(output);// escribo todo lo que se ha hecho anteriormente
        output.close();
    }

    public static void leerXLSX() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream("LeerExcel.xlsx");

        XSSFWorkbook wb = new XSSFWorkbook(file);

        XSSFSheet sheet = wb.getSheetAt(0);
        int numFilas = sheet.getLastRowNum();
        System.out.println("Leyendo fichero: xlsx");
        for (int a = 0; a <= numFilas; a++) {
            Row fila = sheet.getRow(a);
            int numCols = fila.getLastCellNum();
            for (int b = 0; b < numCols; b++) {
                Cell celda = fila.getCell(b);

                switch (celda.getCellType().toString()) {
                    case "NUMERIC":
                        System.out.print(celda.getNumericCellValue() + " ");
                        break;
                    case "STRING":
                        System.out.print(celda.getStringCellValue() + " ");
                        break;
                    case "FORMULA":
                        System.out.print(celda.getCellFormula() + " ");
                        break;
                }
            }
            System.out.println("");
        }
    }

    public static void ModificaXLSX() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream("LeerExcel.xlsx");

        XSSFWorkbook wb = new XSSFWorkbook(file);

        XSSFSheet sheet = wb.getSheetAt(0);

        // comprobar que exista la fila y la celda...
        XSSFRow fila = sheet.getRow(1);

        if (fila == null) {
            fila = sheet.createRow(1);

        }

        XSSFCell celda = fila.getCell(1); //leer la celda para conservar formato.
        if (celda == null) {
            celda = fila.createCell(1);
        }

        celda.setCellValue("Modificacion");

        file.close();

        FileOutputStream output = new FileOutputStream("Modificado.xlsx");
        wb.write(output);// escribo todo lo que se ha hecho anteriormente
        output.close();
    }
}
