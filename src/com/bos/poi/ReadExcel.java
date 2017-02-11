package com.bos.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

/**
 * ReadExcel
 * @author 山长鲁
 * @email  changlu1119@gmail.com
 * @time   2017年2月3日 下午5:08:15 
 * @version 1.0
 */
public class ReadExcel {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		/**
		 * 通过指定的EXCEL文件创建工作簿
		 */
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("poi2.xls"));
		System.out.println("工作单的总数量"+workbook.getNumberOfSheets());
		System.out.println("第一个工作单的名字"+workbook.getSheetName(0));
		/**
		 * 通过工作簿获取工作单
		 */
		HSSFSheet sheet = workbook.getSheetAt(0);
		/**
		 * 通过工作单获取所有的行
		 */
		Iterator<Row> rows = sheet.rowIterator();
		/**
		 * 迭代所有行
		 */
		while (rows.hasNext()) {
			/** 获取一行 */
			Row row = rows.next();
			/** 获取行的列 */
			Iterator<Cell> cells = row.cellIterator();
			/** 迭代所有列 */
			while (cells.hasNext()) {
				/** 获取一列 */
				Cell cell = cells.next();
				/** 获取列中的值 */
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {//数值|日期
					/** 判断是不是日期 */
					if (DateUtil.isCellDateFormatted(cell)) {
						//能被格式化就是日期
						Date date = cell.getDateCellValue();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
						System.out.print(sdf.format(date)+"\t");
					}else {
						System.out.print(cell.getNumericCellValue()+"\t");
					}
				}else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {//布尔
					System.out.print(cell.getBooleanCellValue()+"\t");
				} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {//字符串
					System.out.print(cell.getStringCellValue()+"\t");
				}
			}
			System.out.println();
		}
		/** 关闭工作簿 */
		workbook.close();
	}
}
