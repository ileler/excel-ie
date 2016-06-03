package org.ileler.excel.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class ExcelUtil {
	
	public static final String DATEPATTERN = "yyyy-MM-dd HH:mm:ss";
	
	private ExcelUtil() {};
	
	public static class Config {
		
		private String datePattern;	//按日期格式读取列数据时的格式模板
		private boolean isAutoTrim = true;	//是否自动去掉空格
		private boolean isAllToString = true;	//是否全部按文本格式读取数据

		public String getDatePattern() {
			return !StringUtils.isEmpty(datePattern) ? datePattern : DATEPATTERN;
		}

		/**
		 * 按日期格式读取列数据时的格式模板，默认"yyyy-MM-dd HH:mm:ss"
		 * @param datePattern
		 * @return
		 */
		public Config setDatePattern(String datePattern) {
			this.datePattern = datePattern;
			return this;
		}

		public boolean isAutoTrim() {
			return isAutoTrim;
		}

		/**
		 * 是否自动去掉空格，默认true
		 * @param isAutoTrim
		 * @return
		 */
		public Config setAutoTrim(boolean isAutoTrim) {
			this.isAutoTrim = isAutoTrim;
			return this;
		}

		public boolean isAllToString() {
			return isAllToString;
		}

		/**
		 * 是否全部按文本格式读取数据，默认true
		 * @param isAllToString
		 * @return
		 */
		public Config setAllToString(boolean isAllToString) {
			this.isAllToString = isAllToString;
			return this;
		}
		
	}

	public static Map<Integer, String[]> read(String path) {
		return read(path, 0);
	}
	
	public static Map<Integer, String[]> read(String path, int sheetIndex) {
		return read(path, sheetIndex, new Config());
	}

	public static Map<Integer, String[]> read(String path, Config config) {
		return read(path, 0, config);
	}

	public static Map<Integer, String[]> read(String path, int sheetIndex, Config config) {
		if (StringUtils.isEmpty(path)) 	return null;
		try (InputStream is = new FileInputStream(path);) {
			return read(is, getFilenameExtension(path), sheetIndex, config);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	private static Map<Integer, String[]> read(InputStream is, String suffix, int sheetIndex, Config config) {
		if (is == null || StringUtils.isEmpty(suffix)) 	return null;
		if (config == null) 	config = new Config();
		return "xlsx".equalsIgnoreCase(suffix) ? readXLSX(is, sheetIndex, config) : readXLS(is, sheetIndex, config);
	}
	
	public static Map<Integer, String[]> readXLS(InputStream is) {
		return readXLS(is, 0);
	}

	public static Map<Integer, String[]> readXLSX(InputStream is) {
		return readXLSX(is, 0);
	}
	
	public static Map<Integer, String[]> readXLS(InputStream is, int sheetIndex) {
		return readXLS(is, sheetIndex, new Config());
	}

	public static Map<Integer, String[]> readXLSX(InputStream is, int sheetIndex) {
		return readXLSX(is, sheetIndex, new Config());
	}

	public static Map<Integer, String[]> readXLS(InputStream is, Config config) {
		return readXLS(is, 0, config);
	}

	public static Map<Integer, String[]> readXLSX(InputStream is, Config config) {
		return readXLSX(is, 0, config);
	}
	
	public static Map<Integer, String[]> readXLS(InputStream is, int sheetIndex, Config config) {
		try {
			return read(new HSSFWorkbook(new POIFSFileSystem(is)).getSheetAt(sheetIndex), config);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	public static Map<Integer, String[]> readXLSX(InputStream is, int sheetIndex, Config config) {
		try {
			return read(new XSSFWorkbook(is).getSheetAt(sheetIndex), config);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	/**
	 * 读取Excel的Sheet里面的数据(一行所有列组成一个String数组)
	 * @param sheet
	 * @param config
	 * @return
	 */
	private static Map<Integer, String[]> read(Sheet sheet, Config config) {
		if (sheet == null) 	return null;
		int rowNum = sheet.getLastRowNum(); 	//得到总行数
		Map<Integer, String[]> content = new HashMap<Integer, String[]>();
		for (int i = 0; i <= rowNum; i++) {
			try {
				Row row = sheet.getRow(i);
				String[] str = null;
				if (row != null) {
					int colNum = row.getLastCellNum();  //得到总列数
					if (colNum < 1)     continue;
					str = new String[colNum];
					int j = 0;
					while (j < colNum) {
						str[j] = getCellFormatValue(row.getCell(j++), config);
					}
				}
				content.put(i, str);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return content;
	}
	
	private static String getFilenameExtension(String path) {
		int i = -1;
		if (StringUtils.isEmpty(path) || (i = path.lastIndexOf(".")) == -1) 	return "";
		return path.substring(i + 1);
	}
	
	/**
	 * 根据配置获取列数据
	 * @param cell
	 * @param config
	 * @return
	 */
	private static String getCellFormatValue(Cell cell, Config config) {
		if (cell == null || config == null) 	return null;
		String cellValue = "";
		if (config.isAllToString()) {
			//统一按文本格式读取数据
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cellValue = cell.getRichStringCellValue().getString();
		} else {
			//判断当前Cell的Type
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC: //如果当前Cell的Type为NUMERIC
				case Cell.CELL_TYPE_FORMULA:
					if (DateUtil.isCellDateFormatted(cell)) { //判断当前的cell是否为Date
						cellValue = new SimpleDateFormat(config.getDatePattern()).format(cell.getDateCellValue());
					} else { //取得当前Cell的数值
						cellValue = String.valueOf(cell.getNumericCellValue());
					}
					break;
				case Cell.CELL_TYPE_STRING: //如果当前Cell的Type为STRIN
					//取得当前的Cell字符串
					cellValue = cell.getRichStringCellValue().getString();
					break;
				default: //默认的Cell值
					cellValue = " ";
			}
		}
		return config.isAutoTrim() ? cellValue.trim() : cellValue;
	}

}
