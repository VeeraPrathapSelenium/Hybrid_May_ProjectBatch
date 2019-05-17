package com.excelplugin;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Method;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.functionlibrary.FunctionLibrary;

public class ReadExcelData {

	public XSSFWorkbook workbook;

	public String crnttestcase;

	public String iteration;

	public boolean intializeDataFile() {
		boolean status = true;

		try {

			System.out.println(System.getProperty("user.dir"));

			String filepath = System.getProperty("user.dir") + "\\TestData\\Testdata.xlsx";

			File excelfile = new File(filepath);

			FileInputStream fis = new FileInputStream(excelfile);

			workbook = new XSSFWorkbook(fis);

		} catch (Exception e) {
			status = false;
			System.out.println("Unable to load data file");
		}

		return status;
	}

	public void executeTestCases() {
		try {

			String testcase = "";
			XSSFSheet sheet = workbook.getSheet("BuisnessFlow");

			int rowcount = sheet.getPhysicalNumberOfRows();

			for (int r = 1; r <= rowcount - 1; r++) {

				// get the cell and verify if it is marked as Yes or No

				String execute_status = sheet.getRow(r).getCell(1).getStringCellValue();

				if (execute_status.toLowerCase().trim().equals("yes")) {
					testcase = sheet.getRow(r).getCell(0).getStringCellValue();
					crnttestcase = testcase;

					System.out.println("Test Case Name :" + testcase);
					int colcount = sheet.getRow(r).getPhysicalNumberOfCells() - 1;
					// read each keyword
					for (int c = 2; c <= colcount; c++) {
						String crntkeword = sheet.getRow(r).getCell(c).getStringCellValue();

						// if keyword is empty then break the loop
						if (!crntkeword.isEmpty()) {
							System.out.println(crntkeword);

							Class oclass = Class.forName("com.functionlibrary.FunctionLibrary");
							Object obj = oclass.newInstance();
							Method method = oclass.getDeclaredMethod(crntkeword, null);

							method.invoke(obj, null);

						} else {
							break;
						}

					}

				}

			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}

	/****
	 * Method Name:getRowCount
	 */

	public int getRowCount(String sheetname) {
		XSSFSheet sheet = workbook.getSheet(sheetname);

		return sheet.getPhysicalNumberOfRows();

	}

	/****
	 * Method Name:getRowCount
	 */

	public int getColCount(String sheetname) {
		XSSFSheet sheet = workbook.getSheet(sheetname);
		int row = getRowCount(sheetname);

		return sheet.getRow(row).getPhysicalNumberOfCells();

	}

	/***
	 *
	 *Method:
	 *Purpose: search verticaly to find the testcase	 * 
	 * @param sheetname
	 * @param iteration
	 * @return
	 */
	public int searchTestcase(String sheetname,int iteration)
	{
		int rowfound=0;
		
		int rowcount=getRowCount(sheetname);
		
		for(int r=1;r<=rowcount;r++)
		{
			String testcase=workbook.getSheet(sheetname).getRow(r).getCell(0).getStringCellValue();
			String itr=String.valueOf(workbook.getSheet(sheetname).getRow(r).getCell(1).getNumericCellValue());
			
			if(testcase.trim().equals(crnttestcase) && itr.equals(String.valueOf(iteration)))
			{
				rowfound=r;
				break;
			}
			
			
		}
		return iteration;
		
	}
	
	
	/***
	 *
	 *Method:
	 *Purpose: search verticaly to find the testcase	 * 
	 * @param sheetname
	 * @param iteration
	 * @return
	 */
	public int searchcolumn(String sheetname,String columnname)
	{
		int colfound=0;
		
		int colcount=getColCount(sheetname);
		
		for(int c=0;c<=colcount;c++)
		{
			String colmn=workbook.getSheet(sheetname).getRow(0).getCell(c).getStringCellValue();
			
			
			if(colmn.trim().equals(columnname))
			{
				colfound=c;
				break;
			}
			
			
		}
		return colfound;
		
	}
	
	
	
	public String getData(String sheetname,String columnname,int iteration)
	{
		int rowpos=searchTestcase(sheetname,iteration);
		int colpos=searchcolumn(sheetname,columnname);
		
		String data="";
		if(!(rowpos==0) && !(colpos==0))
		{
			switch(workbook.getSheet(sheetname).getRow(rowpos).getCell(colpos).getCellTypeEnum())
			{
			
			case STRING:
				data=workbook.getSheet(sheetname).getRow(rowpos).getCell(colpos).getStringCellValue();
				break;
				
			case NUMERIC:
				data=String.valueOf(workbook.getSheet(sheetname).getRow(rowpos).getCell(colpos).getNumericCellValue());
				break;
				
				
				
			}
		}
		return data;
		
	}
	
	
	
	
	
	
	
	
	
}

