package excelOperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	String outPath = "C:\\Users\\M1064642\\eclipse-workspace\\FileIo_Review\\excelOot.txt";

	public void readData() {
		StringBuilder sb = new StringBuilder();
		try {

			String excelFilePath = "C:\\Users\\M1064642\\eclipse-workspace\\FileIo_Review\\myExcel.xlsx";
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet firstSheet = workbook.getSheetAt(0);
			File file = new File(outPath);
			Iterator<Row> iterator = firstSheet.iterator();

			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					int i =0;
					int j =0;
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue());
						sb.append(cell.getStringCellValue() + ", ");
						break;

					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue());
						sb.append(String.valueOf(cell.getNumericCellValue() + ", "));
						break;

					}
					System.out.print("\t");
				}
				System.out.println();
				sb.append("\n");
			}
			workbook.close();
			inputStream.close();
			Path path = Paths.get(outPath);
			Files.write(path, Arrays.asList(sb.toString()));

		} catch (Exception e) {
			e.printStackTrace();
		}
	
	}

public int search(String name) throws IOException {
		
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
						System.out.print(cell.getStringCellValue());
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


}
