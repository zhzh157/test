package cn.itcast.poi.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class PoiTest01 {

	/**
	 * 创建一个excel
	 *      创建excel：
	 *          1.创建工作簿
	 *          2.创建sheet
	 *          3.创建行对象
	 *          4.创建单元格
	 *          5.对单元格赋值
	 *          6.设置样式
	 *          7.下载
	 */
	@Test
	public void test() throws Exception {
		//1.创建一个工作簿
		//Workbook wb = new HSSFWorkbook(); //处理excel2003版本      .xls
		Workbook wb = new XSSFWorkbook();//处理excel2007及以上版本    .xlsx
		//new SXSSFWorkbook();// 处理大数据量excel报表对象
		//2.创建sheet
		Sheet sheet = wb.createSheet("abc");
		//3.创建行对象
		Row row = sheet.createRow(1);//接受参数 ，数组下标
		//4.创建单元格
		Cell cell = row.createCell(1);//数组下表
		//5.设置单元格内容
		cell.setCellValue("传智播客");

		//设置样式
		/**
		 * 1.创建样式对象
		 * 2.通过样式对象指定样式
		 * 3.配置单元个样式
		 */
		CellStyle cellStyle = wb.createCellStyle();
		//通过样式对象指定样式
		cellStyle.setBorderTop(BorderStyle.THIN); //细线
		cellStyle.setBorderBottom(BorderStyle.THIN); //细线
		cellStyle.setBorderLeft(BorderStyle.THIN); //细线
		cellStyle.setBorderRight(BorderStyle.THIN); //细线

		//字体 对象
		Font font = wb.createFont();
		font.setFontName("华文行楷");
		font.setFontHeightInPoints((short)26);//字号

		cellStyle.setFont(font);

		cell.setCellStyle(cellStyle);

		//指定行高和列宽
		sheet.setColumnWidth(1,20*256); //列宽  不准确！！！
		row.setHeightInPoints(30);



		//6.将excel保存到本地磁盘中
		FileOutputStream fos = new FileOutputStream("E:\\text.xlsx");
		wb.write(fos);
		fos.close();
	}
}
