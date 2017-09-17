import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_read {
	private static final String FILE_NAME = "C:\\data_excel\\phanindra_company_list_30_08_2017.xlsx";

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		System.out.println("Hello world");
		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		// Workbook workbook = new XSSFWorkbook(excelFile);
		// Sheet datatypeSheet = workbook.getSheetAt(0);
		// Sheet datatypeSheet1 = workbook.getSheetAt(1);
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		XSSFSheet datatypeSheet = workbook.getSheetAt(0);
		XSSFSheet datatypeSheet1 = workbook.getSheetAt(1);
		System.out.println(datatypeSheet.getRow(2).getCell(2).getStringCellValue());
		System.out.println(datatypeSheet.getPhysicalNumberOfRows());
		System.out.println(datatypeSheet.getLastRowNum());
		XSSFCell cell;
		XSSFRow row;
		for (int i = 1; i <= datatypeSheet.getLastRowNum(); i++) {
			row = datatypeSheet.getRow(i); //get row value
			if (row == null) {//check if row is empty
				System.out.println(i+" row is empty");
			} else {
				//get cell values
				cell = row.getCell(1, row.RETURN_BLANK_AS_NULL);
				if (cell == null) {//check if cell is empty
					System.out.println("("+i+",1)"+"cell is empty");
				} else {
					if (datatypeSheet.getRow(i).getCell(1).getStringCellValue()
							.equals(datatypeSheet1.getRow(i).getCell(0).getStringCellValue())) {
						if (datatypeSheet.getRow(i).getCell(2).getStringCellValue()
								.equals(datatypeSheet1.getRow(i).getCell(3).getStringCellValue())) {
							System.out.println("find");
							//As since the row already exists, no need to use add row. if added it will overwrite entire row.
							datatypeSheet.getRow(i).createCell(7)
									.setCellValue(datatypeSheet1.getRow(i).getCell(1).getNumericCellValue());
						}
					}
				}
			}
		}
		datatypeSheet.createRow(6).createCell(4).setCellValue("phanindraAHS");
		FileOutputStream outputStream = new FileOutputStream(new File(FILE_NAME));

		workbook.write(outputStream);
		outputStream.close();
		System.out.println(datatypeSheet.getRow(6).getCell(4).getStringCellValue());
		excelFile.close();
	}
}
