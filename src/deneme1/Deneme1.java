/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package deneme1;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.sql.SQLException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.*;
import java.util.Scanner;

public class Deneme1 {
   public static final String URL="jdbc:ucanaccess://C:/Database22.accdb;}";

    public static  double ürünadedi;
    public static  String ürünadı;
    public static  double kodu;
    public static  Scanner Myscan;
  
    public static void main(String[] args) throws Exception{
        int sayi=0;
           Connection Baglanti=null;
        
         
            Baglanti=DriverManager.getConnection(URL);
            System.out.println("Bağlantı açıldı !");
            Statement s=Baglanti.createStatement();
            int i=0;
         
   try {

     

    FileInputStream file = new FileInputStream(new File("C:\\test20.xls"));

     

    HSSFWorkbook workbook = new HSSFWorkbook(file);


    HSSFSheet sheet = workbook.getSheetAt(0);

    

    Iterator<Row> rowIterator = sheet.iterator();
    
    while(rowIterator.hasNext()) {

        Row row = rowIterator.next();

         

        Iterator<Cell> cellIterator = row.cellIterator();

        while(cellIterator.hasNext()) {

            Cell cell = cellIterator.next();

            switch(cell.getCellType()) {

                case Cell.CELL_TYPE_NUMERIC:
                    sayi++;
                    if(i==0){
                    ürünadedi=cell.getNumericCellValue();
                    i=1;
                    
                    }
                    else{
                    kodu=cell.getNumericCellValue();
                    
                    i=0;
                    }
                    if(sayi==3){
                    
                     String SQL="INSERT INTO Table1(urunadedi,urunadı,urunkodu) VALUES('"+ürünadedi+"','"+ürünadı+"','"+kodu+"')";
       s.executeUpdate(SQL);
       sayi=0;
                    }
                    break;

                case Cell.CELL_TYPE_STRING:
               sayi++;
                    ürünadı=cell.getStringCellValue() ;

                    break;

            }

        }

        System.out.println("");

    }

    file.close();

    FileOutputStream out = 

        new FileOutputStream(new File("C:\\test20.xls"));

  workbook.write(out);

    out.close();

     

} catch (FileNotFoundException e) {

    e.printStackTrace();

} catch (IOException e) {

    e.printStackTrace();

}
        
    }
           
             
        }
        
        
           

        
    

   

