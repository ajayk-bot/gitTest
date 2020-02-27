import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

public class printExcelData {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream("E:\\Information\\ExcelDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int numberOfSheet = workbook.getNumberOfSheets();

		for (int i = 0; i < numberOfSheet; i++) {

			XSSFSheet sheet = workbook.getSheetAt(i);
			// System.out.println(workbook.getSheetName(i));
			int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
			// System.out.println("Each sheet rowcount is "+ rowCount);

			for (int j = 0; j < rowCount + 1 ; j++) {
				XSSFRow row = sheet.getRow(j);
				//System.out.println(row.getLastCellNum());
				for (int k = 0; k < row.getLastCellNum(); k++) {
					if (row.getCell(k).getCellType() == CellType.STRING) {
						System.out.println(row.getCell(k).getStringCellValue());
					} else {
						System.out.println(NumberToTextConverter.toText(row.getCell(k).getNumericCellValue()));
						//System.out.println(row.getCell(k).getStringCellValue() + "||");
					}
				}
			}

		}
	}

}
