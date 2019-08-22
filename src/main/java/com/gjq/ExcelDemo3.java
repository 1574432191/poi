package com.gjq;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
/**
 * 读取Excel表格
 * @author Administrator
 *
 */
public class ExcelDemo3 {
	@Test
	public void readExcel() throws IOException{
		//把数据读到内存中
		FileInputStream inputStream = 
				new FileInputStream(new File("C:\\demo.xls"));
		//读取工作簿
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		//读取工作簿，或者getSheet("hello")
		HSSFSheet sheet = workbook.getSheetAt(0);
		//读取行
		for (int i = 1; i <=2; i++) {
			HSSFRow row = sheet.getRow(i);
			//读取单元格，并读值
			HSSFCell cell0 = row.getCell(0);
			String value0 = cell0.getStringCellValue();
			//获取单元格，并读值
			HSSFCell cell1 = row.getCell(1);
			String value1 = cell1.getStringCellValue();
			System.out.println(value0+" + "+value1);
		}
		//关闭流
		inputStream.close();
		//最后记得关闭工作簿
		workbook.close();
	}

}
