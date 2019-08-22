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

public class ExcelDemo1 {
	@Test
	public void writeExcel() throws FileNotFoundException, IOException{
		//创建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		//创建工作表
		HSSFSheet sheet = workbook.createSheet("hello");
		
		//创建行 ，第一行
		HSSFRow row1 = sheet.createRow(0);
		//创建列，1行第一列，设置第一行1列的值
		HSSFCell cell10 = row1.createCell(0);
		cell10.setCellValue("姓名");
		//创建列，1行第2列，设置第一行2列的值
	    HSSFCell cell11 = row1.createCell(1);
		cell11.setCellValue("年龄");
		
		//创建第二行
		HSSFRow row2 = sheet.createRow(1);
		//创建第二行第一列，并设置第二行第一列的值
		HSSFCell cell20 = row2.createCell(0);
		cell20.setCellValue("张三");
		//创建第二行第2列，并设置第二行第二列的值
		HSSFCell cell21 = row2.createCell(1);
		cell21.setCellValue("18");
		
		//创建第3行
		HSSFRow row3 = sheet.createRow(2);
		//创建第3行第1列，并设置第3行第1列的值
		HSSFCell cell30 = row3.createCell(0);
		cell30.setCellValue("李四");
		//创建第3行第2列，并设置第3行第2列的值
		HSSFCell cell31 = row3.createCell(1);
		cell31.setCellValue("20");
		//生成文件
		workbook.write(new FileOutputStream(new File("C:\\demo.xls")));
		//最后记得关闭工作簿
		workbook.close();
		
	}

}
