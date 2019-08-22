package com.gjq;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

public class ExcelDemo2 {
	@Test
	public void writeExcel() throws FileNotFoundException, IOException{
		//创建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		//创建工作表，工作表（sheet）的名字叫hello
		HSSFSheet sheet = workbook.createSheet("hello");
		
		//创建行，第一行
		HSSFRow row1 = sheet.createRow(0);
		//创建列，1行第1列，设置一行一列的值
		HSSFCell cell10 = row1.createCell(0);
		cell10.setCellValue("姓名");
		//创建列，第一行第二列，设置1行2列的值
		HSSFCell cell11 = row1.createCell(1);
		cell11.setCellValue("年龄");
		
		//第一行结束
		for (int i = 1; i <=4; i++) {
			HSSFRow row2 = sheet.createRow(i);
			//创建第2行第1列，并设置第2行第1列的值
			HSSFCell cell21 = row2.createCell(0);
			cell21.setCellValue("张三"+i);
			//创建第2行第2列，并设置第2行第2列的值
			HSSFCell cell22 = row2.createCell(1);
			cell22.setCellValue(18+i);
		}
		//生成文件
		workbook.write(new FileOutputStream(new File("C:\\demo1.xls")));
		//最后关闭工作簿
		workbook.close();
		
	}

}
