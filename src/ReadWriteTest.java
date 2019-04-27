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
		XSSFCell cell = sheet.getRow(8).getCell(2);
		String c9 = cell.getStringCellValue();
		System.out.println(c9);
		
		
		
		
		
		
	}

}
