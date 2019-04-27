import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteTest {

	public static void main(String[] args)  throws Exception{
		File Rsrc = new File("C:\\Users\\mahmoud.mohamed\\Documents\\TLR.xlsx");
		FileInputStream Rfile =  new FileInputStream(Rsrc);
		
		File Wsrc = new File("C:\\Users\\mahmoud.mohamed\\Documents\\TLW.xlsx");
		FileInputStream Wfile =  new FileInputStream(Wsrc);
		
		XSSFWorkbook workbook = new XSSFWorkbook(Rfile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook2 = new XSSFWorkbook(Wfile);
		XSSFSheet sheet2 = workbook2.getSheetAt(0);
		
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
		for (int i = 2; i < count; i++) {
			// printing multy cells 
 			String data = sheet.getRow(8).getCell(i).getStringCellValue();
		//	System.out.println(data);
/*			for(int j=0;j<count;j++) {

			}*/
			
		}
		XSSFCell cellW = sheet2.getRow(9).getCell(2);
		String cellVal = cellW.getStringCellValue();
	//	sheet2.getRow(j).getCell(10).setCellValue(data);
		System.out.println(cellVal);
		
		XSSFCell cellT = sheet2.getRow(9).getCell(2);
		cellT.setCellValue("Mahmoud");
		System.out.println(cellT);
		System.out.println();
		workbook.close();
	}

}
