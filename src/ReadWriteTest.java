import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteTest {

	public static void main(String[] args)  throws Exception{
		File Rsrc = new File("C:\\Users\\mahmoud.mohamed\\Documents\\TLR.xlsx");
		FileInputStream Rfile =  new FileInputStream(Rsrc);
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(Rfile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		// declaring and printing cell C9
		XSSFCell cellc9 = sheet.getRow(8).getCell(2);
		String c9 = cellc9.getStringCellValue();
//		System.out.println(c9);
		
		// declaring and printing cell F9
		XSSFCell cellf9 = sheet.getRow(8).getCell(5);
		String f9 = cellf9.getStringCellValue();
	//	System.out.println(f9);
		
		// declaring and printing cell I9
		XSSFCell cellI9 = sheet.getRow(8).getCell(8);
		String I9 = cellI9.getStringCellValue();
	//	System.out.println(I9);
		
		// iterating through the cells 
		int count = 30;
		for (int i = 0; i < count; i++) {
			// printing multy cells 
			String data = sheet.getRow(8).getCell(i+2).getStringCellValue();
			System.out.println(data);
		}
		System.out.println();
		workbook.close();
	}

}
