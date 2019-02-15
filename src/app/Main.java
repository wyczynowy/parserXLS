package app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Main {
	
	private static String pathToFile = "C:\\Users\\dkozub\\Desktop\\Dokumenty\\xls\\";
	private static String fileName = "Dokument.xls";
	private static String inputFileName = pathToFile + fileName;
	private static String outputFileName = pathToFile + "PO_ZMIANACH\\" + fileName;
	static // Odnajduje ciag znakow pasujacy do wzorca: ${K[1-4 cyfry]}
	Pattern compiledPattern = Pattern.compile("\\$\\{K\\d{1,4}\\}");

	public static void main(String[] args) {

		try {
			File file = new File(inputFileName);
		    POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			
			for(int idx = 0; idx < wb.getNumberOfSheets(); idx++) {
				HSSFSheet sheet = wb.getSheetAt(idx);
			    HSSFRow row;
			    HSSFCell cell;

			    int rows; // No of rows
			    rows = sheet.getPhysicalNumberOfRows();

			    int cols = 0; // No of columns
			    int tmp = 0;

			    // This trick ensures that we get the data properly even if it doesn't start from first few rows
			    for(int i = 0; i < 10 || i < rows; i++) {
			        row = sheet.getRow(i);
			        if(row != null) {
			            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
			            if(tmp > cols) cols = tmp;
			        }
			    }

			    for(int r = 0; r < rows; r++) {
			        row = sheet.getRow(r);
			        if(row != null) {
			            for(int c = 0; c < cols; c++) {
			                cell = row.getCell((short)c);
			                if(cell != null) {
			                	if(cell.getCellType().toString().equals("STRING")) {
			                		Matcher matcher = compiledPattern.matcher(cell.getStringCellValue());
			                		if(matcher.find()) {
			                			/********** TUTAJ DEFINIUJEMY ZMIANY DLA KOMOREK **********/
			                			String cellAddress = cell.getAddress().toString() + ": " + cell.getStringCellValue();
			                			int cellValue = Integer.parseInt(cell.getStringCellValue().substring(3,cell.getStringCellValue().length()-1));
			                			int sheetNumber = idx;
			                			System.out.println( "SheetNumber:" + sheetNumber + " - " + cellAddress + " - " + cellValue);
				                		if(cellValue >= 1211 && cellValue <= 2418)
				                			cell.setCellValue("$K{" + (cellValue - 2) + "}");
				                		if(cellValue >=2421 && cellValue <= 2525)
				                			cell.setCellValue("${K" + (cellValue - 4) + "}");
				                		/********** TUTAJ DEFINIUJEMY ZMIANY DLA KOMOREK **********/
			                		}
			                	}
			                }
			            }
			        }
			    }			
			}
			
		    File outputFile = new File(outputFileName);
	        FileOutputStream out = new FileOutputStream(outputFile);
	        wb.write(out);
	        out.close();
	        wb.close();	
		    
		} catch(Exception ioe) {
		    ioe.printStackTrace();
		    System.err.println("\nBŁĄD - COŚ POSZŁO NIE TAK !");
		}
		
		System.out.println("\nMODYFIKACJA PLIKU PRZEBIEGŁA POMYŚLNIE !");
		
	}
}
