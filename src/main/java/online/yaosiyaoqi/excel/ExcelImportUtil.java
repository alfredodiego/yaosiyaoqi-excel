package online.yaosiyaoqi.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

public class ExcelImportUtil {

	protected Logger log = LoggerFactory.getLogger( ExcelImportUtil.class );

	/**
	 * 
	 * @param file
	 *            文件
	 * @param SheetNum
	 *            读取第几个sheet页，从0开始
	 * @param ColumnNum
	 *            读取多少列
	 * @param rowNum
	 *            从第几行开始读取 从0开始
	 * @return
	 * @throws IOException
	 */
	public static List<String[]> readExcel( File file, int SheetNum, int columnNum, int rowNum ) throws IOException {
		List<String[]> dataList = new ArrayList<String[]>();

		/*
		 * InputStream input = new FileInputStream(file); POIFSFileSystem fs =
		 * new POIFSFileSystem(input); HSSFWorkbook wb = new HSSFWorkbook(fs);
		 * HSSFSheet sheet = wb.getSheetAt(SheetNum); Iterator rows =
		 * sheet.rowIterator(); //逐行读取 int i = 0; while (rows.hasNext()) {
		 * HSSFRow row = (HSSFRow) rows.next(); if(i < rowNum){//从rowNum行开始读取
		 * i++; }else{ String dataArray[] = new String [columnNum]; for (int j =
		 * 0; j < columnNum; j++) { HSSFCell cell = null; cell =
		 * (HSSFCell)row.getCell(j); String value = getCellValue(cell);
		 * dataArray[j] = value; } dataList.add(dataArray); } }
		 */
		List<List<LinkedList<String>>> sheetLists = readExcelData( file );
		if( null!=sheetLists ) {
			List<LinkedList<String>> sheetList = sheetLists.get( 0 );// 数据位于sheet0
			sheetList.remove( 0 );
			for( LinkedList<String> linkedList : sheetList ) {
				String dataArray[] = new String[ linkedList.size() ];
				for( int j = 0; j < linkedList.size(); j++ ) {
					dataArray[ j ] = linkedList.get( j );
				}
				dataList.add( dataArray );
			}
		}
		return dataList;
	}

	// 获得cell中的内容
	private static String getCellValue( HSSFCell cell ) {
		if( cell == null ) {
			return null;
		} else {
			cell.setCellType( Cell.CELL_TYPE_STRING );
			return cell.getStringCellValue();
		}
	}

	/** 开始从excel读取数据 **/
	public static List<List<LinkedList<String>>> readExcelData( File excelFile ) {
		Workbook workBook = null;
		FileInputStream fis = null;
		try {
			if( excelFile == null || !( isExcel2003( excelFile ) || isExcel2007( excelFile ) ) ) { throw new FileNotFoundException(); }
			fis = new FileInputStream( excelFile );
			if( isExcel2003( excelFile ) ) {
				workBook = new HSSFWorkbook( fis );
			} else {
				workBook = new XSSFWorkbook( fis );
			}
		} catch( Exception e ) {
			e.printStackTrace();
		} finally {
			try {
				if(null!=fis){
					fis.close();
				}
			} catch( IOException e ) {
				//
			}
		}
		List<List<LinkedList<String>>> sheetLists = new ArrayList<List<LinkedList<String>>>();
		Sheet sheet = null;
		if( workBook != null ) {
			int sheetSize = workBook.getNumberOfSheets();
			for( int i = 0; i < sheetSize; i++ ) {
				sheet = workBook.getSheetAt( i );
				String entityName = workBook.getSheetName( i );
				sheetLists.add( readSheetData( sheet, entityName ) );
			}
			return sheetLists;
		}
		return null;
	}

	/**
	 * 
	 * isExcel2003
	 * 
	 * @Title: isExcel2003
	 * @Description: 判断是否是2003版excel文件
	 * @param @param file
	 * @param @return 设定文件
	 * @return boolean 返回类型
	 * @throws
	 */
	public static boolean isExcel2003( File file ) {
		String filePath = null;
		if( file != null ) {
			filePath = file.getName();
		} else {
			return false;
		}
		return filePath.matches( "^.+\\.(?i)(xls)$" );
	}

	/**
	 * 
	 * isExcel2007
	 * 
	 * @Title: isExcel2007
	 * @Description: 判断是否是2007版excel文件
	 * @param @param file
	 * @param @return 设定文件
	 * @return boolean 返回类型
	 * @throws
	 */
	public static boolean isExcel2007( File file ) {
		String filePath = null;
		if( file != null ) {
			filePath = file.getName();
		} else {
			return false;
		}
		return filePath.matches( "^.+\\.(?i)(xlsx)$" );
	}

	/** 读每个sheet页的数据 **/
	private static List<LinkedList<String>> readSheetData(Sheet sheet, String entityName ) {
		int rowNumbers = sheet.getPhysicalNumberOfRows();
		if( rowNumbers == 0 ) {
			System.out.println( "================excel中数据为空！" );
			return null;
		}
		int startRow = 0;
		int excelLastRow = sheet.getRow( startRow ).getLastCellNum();
		List<LinkedList<String>> lists = new ArrayList<LinkedList<String>>();
		LinkedList<String> list = null;
		Row columnRow = null;
		for( int i = startRow; i < rowNumbers; i++ ) {
			columnRow = sheet.getRow( i );
			if( columnRow != null ) {
				Cell cell = null;
				list = new LinkedList<String>();
				for( int j = 0; j < excelLastRow; j++ ) {
					cell = columnRow.getCell( j );
					list.add( getStringCellValue( cell ) );
				}
				lists.add( list );
			}
		}
		return lists;
	}

	/**
	 * 获得单元格字符串
	 * 
	 * @throws UnSupportedCellTypeException
	 */
	private static String getStringCellValue( Cell cell ) {
		if( cell == null ) { return null; }

		String result = "";
		switch( cell.getCellType() ) {
		case Cell.CELL_TYPE_BOOLEAN:
			result = String.valueOf( cell.getBooleanCellValue() );
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if( DateUtil.isCellDateFormatted( cell ) ) {
				java.text.SimpleDateFormat TIME_FORMATTER = new java.text.SimpleDateFormat( "yyyy-MM-dd" );
				result = TIME_FORMATTER.format( cell.getDateCellValue() );
			} else {
				double doubleValue = cell.getNumericCellValue();
				result = "" + doubleValue;
			}
			break;
		case Cell.CELL_TYPE_STRING:
			if( cell.getRichStringCellValue() == null ) {
				result = null;
			} else {
				result = cell.getRichStringCellValue().getString();
			}
			break;
		case Cell.CELL_TYPE_BLANK:
			result = null;
			break;
		case Cell.CELL_TYPE_FORMULA:
			try {
				result = String.valueOf( cell.getNumericCellValue() );
			} catch( Exception e ) {
				result = cell.getRichStringCellValue().getString();
			}
			break;
		default:
			result = "";
		}
		return result;
	}

	public static void main( String[] args ) {
		File file = new File( "D:\\poiTest.xls" );
		try {
			List<String[]> dataList = readExcel( file, 0, 2, 1 );
			System.out.println( "==========================" + dataList );
		} catch( IOException e ) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
