import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.print.DocFlavor.CHAR_ARRAY;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadWriteTest {

	public static void main(String[] args) throws IOException  {
		File Rsrc = new File("C:\\Users\\mahmoud.mohamed\\Documents\\TLR.xlsx");
		FileInputStream Rfile =  new FileInputStream(Rsrc);
		
		
		XSSFWorkbook workbook = new XSSFWorkbook(Rfile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		
		// declaring and printing cell C9
		XSSFCell cellc9 = sheet.getRow(8).getCell(2);
		String c9 = cellc9.getStringCellValue();
		System.out.println("C9 contains " + c9);
		
		// declaring and printing cell F9
		XSSFCell cellf9 = sheet.getRow(8).getCell(5);
		String f9 = cellf9.getStringCellValue();
		System.out.println("F9 contains " +f9);
		
		// declaring and printing cell I9
		XSSFCell cellI9 = sheet.getRow(8).getCell(8);
		String I9 = cellI9.getStringCellValue();
		System.out.println("I9 contains " +I9);
		System.out.println();
			
		// iterating through the cells 
		
		// creating new workbook for the data
		@SuppressWarnings("resource")
		HSSFWorkbook WriteOn = new HSSFWorkbook();
		try(OutputStream fileOut = new FileOutputStream("C:\\Users\\mahmoud.mohamed\\Documents\\TLW2.xls")) {
			WriteOn.write(fileOut);
		//	XSSFSheet sheetWrite = WriteOn.createSheet("TLW2");
		}
		
					
		
		int count = 29;
		for (int i = 2; i < count; i++) {
			// printing multy cells 
 			String data = sheet.getRow(8).getCell(i).getStringCellValue();
 			System.out.println(data);
			
		}
		workbook.close();
	}

}
