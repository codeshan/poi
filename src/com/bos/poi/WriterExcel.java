package com.bos.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

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
public class WriterExcel {
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
		 * 通过工作单创建行
		 */
		HSSFRow row = sheet.createRow(0);
		/**
		 * 通过行创建列
		 * 设置列中的文本
		 */
		row.createCell(0).setCellValue("第一列");
		row.createCell(1).setCellValue("第二列");
		row.createCell(2).setCellValue("第三列");
		/**
		 * 把工作簿写入一个输出流
		 */
		workbook.write(new FileOutputStream("poi.xls"));
		/**
		 * 关闭工作簿
		 */
		workbook.close();
	}
}
