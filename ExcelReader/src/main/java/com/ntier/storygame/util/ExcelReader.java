package com.ntier.storygame.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	private static String[] titles= {"id","itemDesc","bv","ec","pr","analysis","code","test","workSum","value","cost","net","vs","avgRoll","depens"};
	
	/*
	 * This static method will read from an excel file with a specific format to generate an ArrayList 
	 * of HashMaps that contain the data from the excel sheet with preset string keys. This allows for
	 * an easier time looping through the data to generate game cards for the agile story game.
	 */
	
	public static ArrayList<HashMap<String,Integer>> getExcelValues(String excelPath,int sheetNum) throws IOException{
		// Creating a xlsx file object with specific file path to read
        File xlsxFile = new File(excelPath);
        ArrayList<HashMap<String,Integer>> rows = new ArrayList<>();
  
        // Creating input stream
        FileInputStream inputStream = new FileInputStream(xlsxFile);
  
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
          
        // Reading the specified sheet of the excel file
        XSSFSheet sheet = workbook.getSheetAt(sheetNum);
          
        Iterator<Row> iterator = sheet.iterator();
        iterator.next();
        iterator.next();
        iterator.next();
        iterator.next();
        
        // Iterating all the rows
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
          
            HashMap<String,Integer> hash = new HashMap<>();
            // Iterating all the columns in a row
            int i =0;
            while (i<15) {
                Cell cell = nextRow.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                switch (cell.getCellType()) {
                case NUMERIC:
                    hash.put(titles[i],(int) cell.getNumericCellValue());
                    break;
                case FORMULA:
                	switch(cell.getCachedFormulaResultType()) {
                    case NUMERIC:
                    	hash.put(titles[i],(int) cell.getNumericCellValue());//reading data from cells with formula
                        break;
					default:
						hash.put(titles[i],-1);
						break;
                	}
                	break;
                default:
                	hash.put(titles[i],-1);
                }
                i++;
                
            }
            rows.add(hash);
        }
  
        // Closing the workbook and input stream
        workbook.close();
        inputStream.close();
        return rows;
	}
	
	

}
