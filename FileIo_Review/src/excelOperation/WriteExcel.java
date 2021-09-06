package excelOperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public int writeData(String name, String updatedCity) throws IOException {
		
		String excelFilePath = "C:\\Users\\M1064642\\eclipse-workspace\\FileIo_Review\\myExcel.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet firstSheet = workbook.getSheetAt(0);
	
		Iterator<Row> iterator = firstSheet.iterator();

		boolean update = false;
		int result = 0;
		
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					//System.out.print(cell.getStringCellValue());
					if(update) {
						cell.setCellValue(updatedCity);
						result = 1;
						update = false;
					}
					
					if(name.equalsIgnoreCase(cell.getStringCellValue()))
					{
						update = true;
					}
					break;

				case Cell.CELL_TYPE_NUMERIC:
					break;

				}
		
			}
		
		}
		
		 FileOutputStream output_file =new FileOutputStream(new File(excelFilePath));
		 workbook.write(output_file);
	        output_file.close();
		return result;

	}
	
	
	
public int addString() throws IOException {
		
		String excelFilePath = "C:\\Users\\M1064642\\eclipse-workspace\\FileIo_Review\\myExcel.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet firstSheet = workbook.getSheetAt(0);
	
		Iterator<Row> iterator = firstSheet.iterator();

		int result = 0;
		int m =0;
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			int p =0;
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					//System.out.print(cell.getStringCellValue());
					if(p == 2 && m !=0) {
						cell.setCellValue(cell.getStringCellValue() + " City");
						result = 1;
					}
					
					break;

				case Cell.CELL_TYPE_NUMERIC:
					break;

				}
				p = p+1;
			}
		m= m+1;
		}
		
		
		 FileOutputStream output_file =new FileOutputStream(new File(excelFilePath));
		 workbook.write(output_file);
	        output_file.close();
		return result;

	}
}
