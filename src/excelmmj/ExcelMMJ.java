/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package excelmmj;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author 52561
 */
public class ExcelMMJ {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        crearExcel();
    }
    
    public static void crearExcel(){
        
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Hola Java");
        
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Nombre");
        row.createCell(1).setCellValue("Apellido");
        row.createCell(2).setCellValue("Numero De Telefono");
        row.createCell(3).setCellValue("Correo");
        row.createCell(4).setCellValue("Ocupacion");
        
        Row rowUno = sheet.createRow(1);
        rowUno.createCell(0).setCellValue("Moises");
        rowUno.createCell(1).setCellValue("Mendez");
        rowUno.createCell(2).setCellValue("5618068815");
        rowUno.createCell(3).setCellValue("moises31102002@gmail.com");
        rowUno.createCell(4).setCellValue("Estudiante");
        
        Row rowDos = sheet.createRow(2);
        rowDos.createCell(0).setCellValue("Mailin");
        rowDos.createCell(1).setCellValue("MCampos");
        rowDos.createCell(2).setCellValue("5541186492");
        rowDos.createCell(3).setCellValue("alailamailin123@gmail.com");
        rowDos.createCell(4).setCellValue("Estudiante");
        
        try {
            FileOutputStream fileout = new FileOutputStream("Excel.xls");
            book.write(fileout);
            fileout.close();
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelMMJ.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelMMJ.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
}
