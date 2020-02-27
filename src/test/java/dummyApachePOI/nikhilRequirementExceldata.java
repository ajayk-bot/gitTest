package dummyApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class nikhilRequirementExceldata {

	public void readingExceldata() throws IOException {
		
		FileInputStream fis = new FileInputStream("E:\\Information\\ExcelDriven1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int no_of_sheets = workbook.getNumberOfSheets();
		System.out.println(no_of_sheets); // print number of sheet
		//print the name of sheet in excel
		for (int i=0; i<no_of_sheets;i++) {
			System.out.println(workbook.getSheetName(i));
		}
		
		for(int i=0 ;i<no_of_sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("LoginTestCase")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstrow =  rows.next();
				Iterator<Cell> cell = firstrow.cellIterator();
				int k=0;
				int coloumn = 0;
				while(cell.hasNext()) {
					Cell value = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						coloumn = k;
					}
					
					k++;
				}
				System.out.println("Testcase present in the coloumn : "+coloumn);
				System.out.println("Name of testcases present in coloumn");
				while(rows.hasNext()) {
					Row row = rows.next();
					//System.out.println(row.getCell(coloumn).getStringCellValue());
					String name_of_testCase = row.getCell(coloumn).getStringCellValue();
					
					System.out.println(name_of_testCase);
				}
				
				
			}
			
		}
		
		
		
	}
	
	public static void main(String[] args) throws IOException {
		nikhilRequirementExceldata data = new nikhilRequirementExceldata();
		data.readingExceldata();
		System.out.println("testing");
	

		
	}

}
