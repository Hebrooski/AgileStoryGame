package com.ntier.storygame.ExcelReader;

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
	
	public static void main(String[] args)  {
		try {
			ArrayList<HashMap<String, Integer>> rows = getExcelValues("C:\\Users\\Grayson\\Downloads\\temp\\WorkshopScenario.xlsx", 2);
			for(HashMap<String,Integer> map :rows) {
				System.out.print("-----------------------------------------------------------------------\n"+
						"|id "+map.get("id")+"|");
				if (map.get("itemDesc")!=-1) 
					System.out.print("                           item decription: "+map.get("itemDesc"));
				else
					System.out.print("                           item decription: ");
				if (map.get("bv")!=-1) 
					System.out.print("    bv: "+map.get("bv"));
				else
					System.out.print("    bv: ");
				if (map.get("ec")!=-1) 
					System.out.print("    ec: "+map.get("ec")+" |\n");
				else
					System.out.print("    ec:   |\n");
				
				System.out.print("-----------------------------------------------------------------------\n\n");
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static ArrayList<HashMap<String,Integer>> getExcelValues(String excelPath,int sheetNum) throws IOException{
		// Creating a xlsx file object with specific file path to read
        File xlsxFile = new File(excelPath);
        ArrayList<HashMap<String,Integer>> rows = new ArrayList<>();
  
        // Creating input stream
        FileInputStream inputStream = new FileInputStream(xlsxFile);
  
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
          
        // Reading the first sheet of the excel file
        XSSFSheet sheet = workbook.getSheetAt(2);
          
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
                    	hash.put(titles[i],(int) cell.getNumericCellValue());
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
