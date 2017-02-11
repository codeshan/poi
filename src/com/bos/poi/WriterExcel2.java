package com.bos.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * WriterExcel
 * @author 山长鲁
 * @email  changlu1119@gmail.com
 * @time   2017年2月3日 下午3:47:48 
 * @version 1.0
 */
public class WriterExcel2 {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		/**
		 * 创建工作簿
		 */
		HSSFWorkbook workbook = new HSSFWorkbook();
		/**
		 * 通过工作簿创建工作单
		 */
		HSSFSheet sheet = workbook.createSheet("first");
		/**
		 * 通过工作单创建行，循环创建
		 */
		for (int i = 0; i < 10; i++) {
			HSSFRow row = sheet.createRow(i);
			/**
			 * 循环创建列
			 */
			for (int j = 0; j < 10; j++) {
				/** 创建一列 */
				HSSFCell cell = row.createCell(j);
				/** 设置列中的值 */
				cell.setCellValue("单元格"+i+j);
				System.out.println(cell);
			}
		}
		/**
		 * 把工作簿写入一个输出流
		 */
		workbook.write(new FileOutputStream("poi2.xls"));
		/**
		 * 关闭工作簿
		 */
		workbook.close();
	}
}
