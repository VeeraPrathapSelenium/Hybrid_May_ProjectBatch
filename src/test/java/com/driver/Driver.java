package com.driver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excelplugin.ReadExcelData;
import com.functionlibrary.FunctionLibrary;

public class Driver {

	public static void main(String[] args) throws IOException, ClassNotFoundException, InstantiationException, IllegalAccessException, NoSuchMethodException, SecurityException, IllegalArgumentException, InvocationTargetException {

		ReadExcelData exl=new ReadExcelData();
		
		boolean status=exl.intializeDataFile();
		
// check if the data file is parsed sucessfully
		
if(status)
{
	exl.executeTestCases();
//excel intialize loop
}

}
}