package com.importedUsers;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App {
	
static Connection crunchifyConn = null;
static PreparedStatement crunchifyPrepareStat = null;

public static void main(String[] arg) {
	
	try {
		makeJDBCConnection(arg[0],arg[1],arg[2]);
		
//		makeJDBCConnection(" jdbc:mysql://localhost:3306/water_system","qusai","root");
		String SAMPLE_XLSX_FILE_PATH = arg[3]; //file name 
//		String SAMPLE_XLSX_FILE_PATH ="new.xlsx";
//		
	    try {
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
			 Sheet sheet = workbook.getSheetAt(0);
			 DataFormatter dataFormatter = new DataFormatter();
			 System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
			 
			 for (int rowNumber = 1; rowNumber <= sheet.getLastRowNum(); rowNumber++) {
			    Row row = sheet.getRow(rowNumber);
			    List<String> cellList = new ArrayList<String>();
			    if (row == null) {
			         // This row is completely empty
			    } else {
			         // The row has data
			    	
			         for (int cellNumber = row.getFirstCellNum(); cellNumber <= row.getLastCellNum(); cellNumber++) {
			             Cell cell = row.getCell(cellNumber);
			             if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			            	 cellList.add("");
			             } else {
			            	 cellList.add(dataFormatter.formatCellValue(cell));
			             }
			             
			         }
			    }
			    addDataToDB(cellList);
			 }
			 
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
		
//		addDataToDB("1",2,"qq","qq","qq",12,"ll",12,13,"12",12,14,"kjlj","sdfsdf",12);
		
		crunchifyPrepareStat.close();
		crunchifyConn.close(); // connection close

	} catch (Exception e) {

		e.printStackTrace();
	}
}

private static void makeJDBCConnection(String URL,String username,String pass) {

	try {
		Class.forName("com.mysql.jdbc.Driver");
		log("Congrats - Seems your MySQL JDBC Driver Registered!");
	} catch (ClassNotFoundException e) {
		e.printStackTrace();
		return;
	}

	try {
		// DriverManager: The basic service for managing a set of JDBC drivers.
		crunchifyConn = DriverManager.getConnection(URL, username, pass);
		if (crunchifyConn != null) {
			log("Connection Successful!");
		} else {
			log("Failed to make connection!");
		}
	} catch (SQLException e) {
		log("MySQL Connection Failed!");
		e.printStackTrace();
		return;
	}

}

private static void addDataToDB(List<String> list) {
	 int stand_to_id = 0;
	 int stand_su_id= 0;
	 int portion= 0;
	 int ward= 0;
	 int cell_no= 0;
	 int unit_sno= 0;
	 int unit_bno= 0;
	 
	 if(list.get(0)!=null&&!list.get(0).isEmpty()) {
		 stand_to_id = Integer.parseInt(list.get(0));
	 }
	 if(list.get(1)!=null&&!list.get(1).isEmpty()) {
		 stand_su_id =Integer.parseInt(list.get(1));
	 }
	 String suname =list.get(2);
	 String st_no =list.get(3);
	 
	 if(list.get(4)!=null&&!list.get(4).isEmpty()) {
		 portion =Integer.parseInt(list.get(4));
	 }
	 if(list.get(5)!=null&&!list.get(5).isEmpty()) {
		 ward =Integer.parseInt(list.get(5));
	 }
	 String	acc_nr =list.get(6);
	 
	 String per_idperinit =list.get(7);
	 
	 String initials =list.get(8);
	 
	 String pername =list.get(9);
	 
	 if(list.get(10)!=null &&!list.get(10).isEmpty()) {
		 cell_no =Integer.parseInt(list.get(10));
	 }
	 
	 String vstreet =list.get(11);
	 
	 if(list.get(12)!=null&&!list.get(12).isEmpty()) {
		 unit_sno =Integer.parseInt(list.get(12));
	 }
	 
	 String vbuild =list.get(13);
	 
	 if(list.get(14)!=null&&!list.get(14).isEmpty()) {
		 unit_bno =Integer.parseInt(list.get(14));
	 }

	try {
		String insertQueryStatement = "INSERT  INTO  imported_users  VALUES  (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

		crunchifyPrepareStat = crunchifyConn.prepareStatement(insertQueryStatement);
		crunchifyPrepareStat.setInt(1,0 );
		crunchifyPrepareStat.setString(2, initials);
		crunchifyPrepareStat.setInt(3, cell_no);
		crunchifyPrepareStat.setString(4, acc_nr);
		crunchifyPrepareStat.setString(5, per_idperinit);
		crunchifyPrepareStat.setString(6, pername);
		crunchifyPrepareStat.setInt(7, portion);
		crunchifyPrepareStat.setString(8, st_no);
		crunchifyPrepareStat.setInt(9, stand_su_id);
		crunchifyPrepareStat.setInt(10, stand_to_id);
		crunchifyPrepareStat.setString(11, suname);
		crunchifyPrepareStat.setInt(12, unit_bno);
		crunchifyPrepareStat.setInt(13, unit_sno);
		crunchifyPrepareStat.setString(14, vbuild);
		crunchifyPrepareStat.setString(15, vstreet);
		crunchifyPrepareStat.setInt(16, ward);
		// execute insert SQL statement
		crunchifyPrepareStat.executeUpdate();
	} catch (

	SQLException e) {
		e.printStackTrace();
	}
}


// Simple log utility
private static void log(String string) {
	System.out.println(string);

}
}