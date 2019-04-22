package cn.itcast.parse;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;

public class ExcelParse {

	/**
	 * 帮助我们进行excel逐行读取，逐行加载
	 * @param path
	 * @throws Exception
	 */
	public void parse (String path) throws Exception {
		//解析器
		SheelHandler hl = new SheelHandler();
		//1.根据 Excel 获取 OPCPackage 对象
		OPCPackage pkg = OPCPackage.open(path, PackageAccess.READ);
		try {
			//2.创建 XSSFReader 对象
			XSSFReader reader = new XSSFReader(pkg);
			//3.获取 SharedStringsTable 对象
			SharedStringsTable sst = reader.getSharedStringsTable();
			//4.获取 StylesTable 对象
			StylesTable styles = reader.getStylesTable();
			XMLReader parser = XMLReaderFactory.createXMLReader();
			// 处理公共属性
			parser.setContentHandler(new XSSFSheetXMLHandler(styles,sst, hl,
					false));
			XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator)
					reader.getSheetsData();

			//逐行读取逐行解析
			while (sheets.hasNext()) {
				InputStream sheetstream = sheets.next();
				InputSource sheetSource = new InputSource(sheetstream);
				try {
					parser.parse(sheetSource);
				} finally {
					sheetstream.close();
				}
			}
		} finally {
			pkg.close();
		}
	}

	public static void main(String[] args) throws Exception {
		new ExcelParse().parse("D:\\测试.xlsx");
	}
}
