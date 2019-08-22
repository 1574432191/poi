package com.gjq;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
/**
 * 画边框线：边框线属于样式范畴，样式范畴的作用是这个工作簿
 * @author Administrator
 *
 */
public class ExcelDemo4 {
	@Test
	public void writeExcelHigh() throws FileNotFoundException, IOException{
		//创建工作簿
				HSSFWorkbook workBook = new HSSFWorkbook();
				// 内容样式的类型是:HSSFCellStyle 
		        HSSFCellStyle style_content = workBook.createCellStyle();
		        style_content.setBorderBottom(HSSFCellStyle.BORDER_THIN);//下边框
		        style_content.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
		        style_content.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
		        style_content.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
				
				//创建工作表  工作表（sheet）的名字叫hello
				HSSFSheet sheet = workBook.createSheet("hello");
				
				//----------------------------------------------------
				//创建行,第1行
				HSSFRow row = sheet.createRow(0);
				//创建列，1行第1列，设置第1行1列的值
				HSSFCell cell0 = row.createCell(0);
				//创建列，1行第2列，设置第1行2列的值
				HSSFCell cell1 = row.createCell(1);
				HSSFCell cell2 = row.createCell(2);
				//设置单元格边框
				cell0.setCellStyle(style_content);
				cell1.setCellStyle(style_content);
				cell2.setCellStyle(style_content);
				//----------------------------------------------------
				//创建行,第1行
				HSSFRow row1 = sheet.createRow(1);
				//创建列，1行第1列，设置第1行1列的值
				HSSFCell cell10 = row1.createCell(0);
				cell10.setCellValue("姓名");
				//创建列，1行第2列，设置第1行2列的值
				HSSFCell cell11 = row1.createCell(1);
				cell11.setCellValue("年龄");
				HSSFCell cell12 = row1.createCell(2);
				cell12.setCellValue("性别");
				
				cell10.setCellStyle(style_content);
				cell11.setCellStyle(style_content);
				cell12.setCellStyle(style_content);
				//第一行结束---------------------------------------------
				for(int i=2;i<=5;i++) {
					HSSFRow row2 = sheet.createRow(i);
					//创建第2行第1列，并设置第2行第1列的值
					HSSFCell cell20 = row2.createCell(0);
					cell20.setCellValue("张三"+i);
					//创建第2行第2列，并设置第2行第2列的值
					HSSFCell cell21 = row2.createCell(1);
					cell21.setCellValue(18+i);
					HSSFCell cell22 = row2.createCell(2);
					cell22.setCellValue("女"+i);
					
					cell20.setCellStyle(style_content);
					cell21.setCellStyle(style_content);
					cell22.setCellStyle(style_content);
				}
				
				HSSFRow row6 = sheet.createRow(6);
				//创建列，1行第1列，设置第1行1列的值
				HSSFCell cell60 = row6.createCell(0);
				//创建列，1行第2列，设置第1行2列的值
				HSSFCell cell61 = row6.createCell(1);
				HSSFCell cell62 = row6.createCell(2);
				//设置单元格边框
				cell60.setCellStyle(style_content);
				cell61.setCellStyle(style_content);
				cell62.setCellStyle(style_content);
				
				//生成文件
				workBook.write(new FileOutputStream(new File("C:\\demo2.xls")) );
				//最后记得关闭工作簿
				workBook.close();
		
	}

}
