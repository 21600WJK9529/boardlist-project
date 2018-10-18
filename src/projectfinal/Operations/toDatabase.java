/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package projectfinal.Operations;

import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.RandomAccessFile;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import projectfinal.LoadingScreen.HomeScreen;
import projectfinal.LoadingScreen.loading2;

/**
 *
 * @author Matthew
 */
public class toDatabase {
    
     
    private String name;
    private String number;
    private String subject;
    private int yearMark;
    private int finalMark;
    private String sym;
    
    //
    private String kpName;
    private String kpNumber;
    private String grade;
    
    private int num = 0;
    private String close = "";
    
    
    public void reading() throws Exception{
        HomeScreen home = new HomeScreen();
        //=============================================================
        //String filename = "publisher.mdb";
        
        String dbURL = "jdbc:ucanaccess://publisher.mdb";//specify the full pathname of the database
       //dbURL+= filename.trim() + ";DriverID=22;READONLY=true}"; 
        String driverName = "net.ucanaccess.jdbc.UcanaccessDriver";
        //home.message.append("Database 1");
       
        
        Class.forName(driverName);
        Connection con = DriverManager.getConnection(dbURL); 
        // create a java.sql.Statement so we can run queries
        Statement s = con.createStatement();
        
        //============================================================      
      
      DatabaseMetaData dbm = con.getMetaData();
      // check if "employee" table is there
      ResultSet tables = dbm.getTables(null, null, "TESTIP2", null);
      if (tables.next()) {
      // Table exists
      //do nothing
      }
       else {
        // Table does not exist
       String create_Table_stmt="create table TESTIP2 (Name     VARCHAR(30), Number VARCHAR(20))";
       s.executeUpdate(create_Table_stmt); // execute the SQL statement
       
      }
      //=================================================================================================================
     
      FileInputStream file = new FileInputStream("Testing.xls");
      HSSFWorkbook workbook = new HSSFWorkbook(file);
      HSSFSheet sheet = workbook.getSheetAt(0);
      
      Row row;
      
      for(int i = 13; i<sheet.getLastRowNum();i++){
          row = (Row) sheet.getRow(i);
          
          if(row.getCell(0)==null){
              this.name = name;
          }else {
              name = row.getCell(0).toString();
              
          }
          
          if(row.getCell(1)==null){
              this.number = number;
          }else {
              number = row.getCell(1).toString();
             
          }
          
          
          
         
          
          s.executeUpdate("insert into TESTIP2 values   ('"+name+"','"+number+"')");
          s.executeUpdate("DELETE FROM TESTIP2 WHERE    Name = ''");
         
      }
      
      //====================================================
      System.out.println("About to close Statement....");
        s.close(); // close the Statement to let the database know we're done with it
        con.close(); // close the Connection to let the database know we're done with it
        System.out.println("Statement closed successfully....");
        //=================================================================
      
      
}
    
    public void entireTable() throws Exception{
        HomeScreen home = new HomeScreen();
         //=============================================================
        //String filename = "publisher.mdb";
       String dbURL = "jdbc:ucanaccess://publisher.mdb";//specify the full pathname of the database
       //dbURL+= filename.trim() + ";DriverID=22;READONLY=true}"; 
        String driverName = "net.ucanaccess.jdbc.UcanaccessDriver";

       
         
        //System.out.println("About to Load the JDBC Driver....");
        Class.forName(driverName);
        //System.out.println("Driver Loaded Successfully....");
        //System.out.println("About to get a connection....");
        Connection con = DriverManager.getConnection(dbURL); 
        //System.out.println("Connection Established Successfully....");
       // create a java.sql.Statement so we can run queries
       // System.out.println("Creating statement Object....");
        Statement s = con.createStatement();
        //System.out.println("Statement object created Successfully....");
        //System.out.println("About to execute SQL stmt....");
        //============================================================      
      
      DatabaseMetaData dbm = con.getMetaData();
      // check if "employee" table is there
      ResultSet tables = dbm.getTables(null, null, "TESTIP1", null);
      if (tables.next()) {
      // Table exists
      
      //replace
      
      }
       else {
        // Table does not exist
             
       String create_Table_stmt2="create table TESTIP1 (Row_num INTEGER,Name VARCHAR(30), Number VARCHAR(10) , Subject VARCHAR(100) , Year_Mark INTEGER, Final_Mark INTEGER, SYM VARCHAR(5) , CLASS VARCHAR(1))";
       s.executeUpdate(create_Table_stmt2); // execute the SQL statement
      }
      //=================================================================================================================
     
      FileInputStream file = new FileInputStream("Testing.xls");
      HSSFWorkbook workbook = new HSSFWorkbook(file);
      HSSFSheet sheet = workbook.getSheetAt(0);
      
      Row row;
      int add = 0;
      int num = 0;
      
      for(int i = 13; i<sheet.getLastRowNum();i++){
          num++;
          row = (Row) sheet.getRow(i);
          
          if(row == null){
             i++; 
          }else{
              
              if(row.getCell(0) == null){
              this.name = name;
          }else {
              name = row.getCell(0).toString().trim();
              
              if(!row.getCell(0).toString().trim().isEmpty()){
                  kpName = name;
              }
              else {
                  name = kpName;
              }
            
          }
          
          if(row.getCell(1)==null){
              this.number = number;
          }else {
              number = row.getCell(1).toString().trim();
              
              if(!row.getCell(1).toString().trim().isEmpty()){
                  kpNumber = number;
              }
              else {
                  number = kpNumber;
              }
             
          }
          
          if(row.getCell(2)==null){
              this.subject = subject;
          }else {
              subject = row.getCell(2).toString().trim();
             
          }
          
          if(row.getCell(3)==null){
              this.yearMark = yearMark;
          }else {
              yearMark = Integer.parseInt(row.getCell(3).toString().trim());
             
          }
          
          if(row.getCell(4)==null){
              this.finalMark = finalMark;
          }else {
              finalMark = Integer.parseInt(row.getCell(4).toString().trim());
             
              if(finalMark == 49){
                  grade = "A";
              }
              else if(finalMark == 48){
                  grade = "B";
              }
              else if(finalMark == 47){
                  grade = "C";
              }
              else if(finalMark == 74){
                  grade = "D";
              }
              else if(finalMark == 73){
                  grade = "E";
              }
              else{
                  grade = "";
              }
              
              
          }
          
          if(row.getCell(5)==null){
              this.sym = sym;
          }else {
              sym = row.getCell(5).toString().trim();
             
          }
          
          }
         
          
          
          s.executeUpdate("insert into TESTIP1 values   ("+num+",'"+name+"','"+number+"' ,'"+subject+"' ,"+yearMark+" ,"+finalMark+" ,'"+sym+"' ,'"+grade+"')");
          //System.out.println("Added to databse...");
          
          
      }
      JOptionPane.showMessageDialog(null, "Database loaded");
      HomeScreen hs = new HomeScreen();
        hs.loading("close");
      
      //====================================================
        //System.out.println("About to close Statement....");
        s.close(); // close the Statement to let the database know we're done with it
        con.close(); // close the Connection to let the database know we're done with it
        
        //System.out.println("Statement closed successfully....");
        //=================================================================
    }
    
    
}
