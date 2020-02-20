import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//Workbook is consider as a whole excel file that gave power to access xls file
//System.out.println("Total Number of sheets in excel : " + workbook.getNumberOfSheets());
//System.out.println(value.getStringCellValue());
//System.out.println("Display Sheets Name in Excel File : "+ workbook.getSheetName(i));
//identify the testCase coloumn in xls file reading each Row to find column name TestCases.
//hasNext method to check whether the immediate naighbour(cell) is present or not..if not written false otherwise true.
//get the firstrow now next to get each cell in the selecte row.

public class dataDriven {

	public ArrayList<String> getData(String testCaseName) throws IOException {

		ArrayList<String> array_list = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("E:\\Information\\ExcelDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int no_of_sheets = workbook.getNumberOfSheets();
		System.out.println(no_of_sheets); // print number of sheet
		//print the name of sheet in excel
		for (int i=0; i<no_of_sheets;i++) {
			System.out.println(workbook.getSheetName(i));
		}

		for (int i = 0; i < no_of_sheets; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("LoginTestCase")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator();
				int k = 0;
				int column = 0;
				while (cell.hasNext()) {
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = k; 
					}
					k++;
				}
				System.out.println(column);

				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							// System.out.println(cv.next().getStringCellValue());
							Cell c = cv.next();
							if(c.getCellType()==CellType.STRING) {
								array_list.add(c.getStringCellValue());
							}else {
								
								array_list.add(NumberToTextConverter.toText((c.getNumericCellValue())));
							}
							
						}

					}
				}
			}

		}

		return array_list;
	}

	public static void main(String[] args) throws IOException {
		dataDriven d1 = new dataDriven();
		//ArrayList listdata = d1.getData("Add Profile");
		//System.out.println(listdata.get(0));
		System.out.println(d1.getData("Add Profile"));
	}

}
