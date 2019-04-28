package Tutorials;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataCofig {
	
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	// path 
	public ExcelDataCofig(String FilePath) {
		try {
			File src = new File(FilePath);
			FileInputStream  fis = new FileInputStream(src);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		} 
	}
	
	public String getData(int sheetNumber, int row, int cell) {
		String data = sheet.getRow(row).getCell(cell).getStringCellValue();
		return data;
	}
}
