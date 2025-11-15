package web.util;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteFile {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("Employee Data");
		Row row;
		row = sheet.createRow(0);
		row.createCell(0).setCellValue("country");
		row.createCell(1).setCellValue("City");
		row.createCell(2).setCellValue("Company");

		for (int i = 1; i <= 7; i++) {
			row = ((XSSFSheet) sheet).createRow(i);

			row.createCell(0).setCellValue("India");
			row.createCell(1).setCellValue("Delhi");
			row.createCell(2).setCellValue("Infosys");

		}

		try (FileOutputStream fileOut = new FileOutputStream("C:\\NewData\\EmployeeData1.xlsx")) {
			workbook.write(fileOut);
			workbook.close();
			System.out.println("Excel file created successfully!");
		}

		catch (IOException e) {
			e.printStackTrace();
		}

	}

}
