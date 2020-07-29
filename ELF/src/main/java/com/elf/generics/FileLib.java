package com.elf.generics;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

	/**
	 * Generic class to read the data from property file or excel
	 * @author shravani
	 */
	public class FileLib {
	/**
	 * 	Used to read the data from property file
	 * @param key
	 * @return the String value of the key
	 * @throws IOException
	 */
	public String getPropertyData(String key) throws IOException {
	FileInputStream fis=new FileInputStream("./data/commondata.property");
	Properties p=new Properties();
	p.load(fis);
	 String value = p.getProperty(key);
	 return value;
	}
	/**
	 * Used to read the data from excel
	 * @param sheetname
	 * @param row
	 * @param cell
	 * @return String value of the excel
 * @throws EncryptedDocumentException
 * @throws IOException
	 * @throws InvalidFormatException 
 */
public String getExcelData(String sheetname,int row,int cell) throws EncryptedDocumentException, IOException, InvalidFormatException {
	FileInputStream fis=new FileInputStream("./data/Test.xlsx");
	Workbook wb = WorkbookFactory.create(fis);
	String data = wb.getSheet(sheetname).getRow(row).getCell(cell).toString();
	return data;
}
/**
 * Used to write the data to excel
 * @param sheetname
 * @param row
 * @param cell
 * @param setvalue
 * @throws EncryptedDocumentException
 * @throws IOException
 * @throws InvalidFormatException 
 */
public void setExcelData(String sheetname,int row,int cell,String setvalue) throws EncryptedDocumentException, IOException, InvalidFormatException {
	FileInputStream fis=new FileInputStream("./data/Test.xlsx");
	Workbook wb = WorkbookFactory.create(fis);
	wb.getSheet(sheetname).getRow(row).getCell(cell).setCellValue(setvalue);
	FileOutputStream fos=new FileOutputStream("./data/Test.xlsx");
	wb.write(fos);
	wb.close();
}
}
