import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
	

	public ArrayList<String>getdata(String testcasename, String Sheetname) throws IOException {
		
		ArrayList<String> a = new ArrayList<String>();

		// Stategy to Access Excel Data:
		// 1. create object for XSSFWorkbook class--- get hold of excel
		// 2. Get Access to desired sheet
		// 3. Get Access to all rows of sheet
		// 4. Access to specific row from all rows
		// 5. Get Access to all cells of Row
		// 6. Access the Data from Excel into Arrays

		// Testing desire:
		// 1. Identify the Testcases column by scanning the entire 1st row
		// 2. once column is identified , then scan entire Testcases column to identify
		// purchase row
		// 3. after you grab purchase Testcase row , pull all the data and feed it into
		// test

		FileInputStream fis = new FileInputStream("src/test/resources/excel data.xlsx");

		//it accepts fileinputstream argument
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase(Sheetname)) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				//1. Identify the Testcases column by scanning the entire 1st row

				Iterator<Row> rows = sheet.iterator();   //sheet is collections of rows

				Row firstrow = rows.next();

				Iterator<Cell> ce = firstrow.cellIterator();   // row is collection of cells

				int k = 0;
				int coloumn = 0;
				while (ce.hasNext()) {

					Cell value = ce.next();

					if (value.getStringCellValue().equalsIgnoreCase(testcasename)) {

						// desired column

						coloumn = k;

					}
					k++;
				}
				
				// once column is identified , then scan entire Testcases column to identify
				// purchase row

				while (rows.hasNext()) 
				{

					Row r = rows.next();

					if (r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename)) {
						
						// after you grab purchase Testcase row , pull all the data and feed it into test

						Iterator<Cell> cv = r.cellIterator();

						while (cv.hasNext()) {
							
						Cell c=	cv.next();
						if (c.getCellType()==CellType.STRING)
							
						{

							a.add(c.getStringCellValue());

						}
						else {
							
							a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
		               }
							
					}
                  
				}

			}

		}
		}
		
		 return a;
	}
	public static void main (String[] args) throws IOException {
		
	}
}
