/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package projectfinal.Operations;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedList;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import projectfinal.LoadingScreen.HomeScreen;

/**
 *
 * @author Matthew
 */
public class txtTOexcel {
    
   static LinkedList<String[]> text_lines = new LinkedList<>();
   static private int endLine = -1;
   
   public void count(String txt) throws FileNotFoundException, IOException{
       BufferedReader br = new BufferedReader(new FileReader(txt));
       
       while(br.readLine() != null){
       endLine++;
       
       
       }
       
   }
    
     public static void open(String txt) throws IOException{
        txtTOexcel text = new txtTOexcel();
        text.count(txt);
        
    try (BufferedReader br = new BufferedReader(new FileReader(txt))) {
        String sCurrentLine;
        int startLine = 0;
        
        int currentLineNo=0;
        while (startLine <= endLine){
    sCurrentLine = br.readLine();
    if(currentLineNo >= startLine && currentLineNo<=endLine) 
    {   
        sCurrentLine = sCurrentLine.replaceAll("                                     ", "\t\t");
        sCurrentLine = sCurrentLine.replaceAll(" {2,}", "\t");
        
        //sCurrentLine = sCurrentLine.replace("(?<=\\s+)1(?=\\s+)", "\t\t\t");
        text_lines.add(sCurrentLine.split("\\t"));
       // System.out.println(sCurrentLine);
        currentLineNo++;
    }else{
        break;
        }
        //while ((sCurrentLine = br.readLine()) != null) {
            //if(sCurrentLine.startsWith(" ")){
              //  text_lines.add(sCurrentLine.split(" "));
            //}else{
               // text_lines.add(sCurrentLine.split("\\t"));
            //}                 
        //}
    } 
    }catch (IOException e) {
        e.printStackTrace();
    }
    } 
    public static void toExcel() throws Exception{
        
        String fileName = "Testing.xls";
        
    Workbook workbook = new HSSFWorkbook();
    Sheet sheet = workbook.createSheet("Test");
    int row_num = 0;
    for(String[] line : text_lines){
        Row row = sheet.createRow(row_num++);
        int cell_num = 0;
        for(String value : line){
            if(value.startsWith("\t")){
               Cell cell = row.createCell(cell_num++ +cell_num++);
                cell.setCellValue(value);  
         
            }else{
                Cell cell = row.createCell(cell_num++);
                cell.setCellValue(value);
            }
           
    }
}
       HomeScreen home = new HomeScreen();
       home.wait.setText("Adding to Database");
       
    try{
       FileOutputStream fileOut = new FileOutputStream(fileName);
       workbook.write(fileOut);
       
       toDatabase data = new toDatabase();
       home.loading("start");
       //data.reading();
       data.entireTable();
       fileOut.close();
        
        
    }catch(Exception e){
        e.getMessage();
    }
        
    }
    
    
}
