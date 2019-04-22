package cn.itcast.poi.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.IOException;

/**
 * 加载excel文件，并读取内容
 */
public class PoiTest02 {

	@Test
	public void test() throws Exception {
		//1.根据excel文件加载工作簿
		Workbook wb = new XSSFWorkbook("E:\\课程资料\\授课文档\\黑马87\\SaaS-Export\\day09\\03-资料\\poi资料\\demo.xlsx");
		//2.读取第一个sheet
		Sheet sheet = wb.getSheetAt(0);//数组下标
		//3.循环sheet中的每一行
		//sheet.getLastRowNum 获取最后一行的数组下标
		for(int i=0;i<sheet.getLastRowNum()+1;i++) {
			Row row = sheet.getRow(i);
			//row.getLastCellNum() 获取最大行数
			//4.读取行中的每一个单元格
			String str = "";
			for(int j=0;j<row.getLastCellNum();j++) {
				Cell cell = row.getCell(j);
				//5.获取单元格中的数据
				if(cell != null) {
					str += getCellValue(cell);
				}
			}
			System.out.println(str);
		}
	}

	public Object getCellValue(Cell cell)  {
		/**
		 * 获取单元格的类型
		 */
		CellType type = cell.getCellType();

		Object result = null;

		switch (type) {
			case STRING:{
				result = cell.getStringCellValue();//获取string类型数据
				break;
			}
			case NUMERIC:{
				/**
				 * 判断
				 */
				if(DateUtil.isCellDateFormatted(cell)) {  //日期格式
					result = cell.getDateCellValue();
				}else{
					//double类型
					result = cell.getNumericCellValue(); //数字类型
				}
				break;
			}
			case BOOLEAN:{
				result = cell.getBooleanCellValue();//获取boolean类型数据
				break;
			}
			default:{
				break;
			}
		}

		return  result;
	}
}
