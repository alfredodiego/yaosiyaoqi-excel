package online.yaosiyaoqi.excel;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * 包含生成excel的样式、常量等相关的通用发放
 */
public class ExcelUtil {

	protected Logger log = LoggerFactory.getLogger( getClass() );

	/**
	 * excel 文件后缀
	 */
	public static final String SUFFIXES2007 = "xlsx";
	public static final String SUFFIXES2003 = "xls";

	/**
	 * 生成sheet页
	 * 
	 * @param sheetNamesArray
	 *            ：每个sheet页面的名称
	 * @return
	 */
	@SuppressWarnings( "deprecation" )
	public static List<HSSFSheet> creatExcel(HSSFWorkbook workbook, String[] sheetNameArray ) {
		List<HSSFSheet> sheetList = new ArrayList<HSSFSheet>();
		for( int i = 0; i < sheetNameArray.length; i++ ) {
			HSSFSheet sheet = workbook.createSheet( sheetNameArray[ i ] );
			// 设置表格默认列宽度为15个字节
			sheet.setDefaultColumnWidth( ( short )25 );
			sheetList.add( sheet );
		}
		return sheetList;
	}

	public static List<XSSFSheet> creatExcel(XSSFWorkbook workbook, String[] sheetNameArray ) {
		List<XSSFSheet> sheetList = new ArrayList<XSSFSheet>();
		for( int i = 0; i < sheetNameArray.length; i++ ) {
			XSSFSheet sheet = workbook.createSheet( sheetNameArray[ i ] );
			// 设置表格默认列宽度为15个字节
			sheet.setDefaultColumnWidth( ( short )25 );
			sheetList.add( sheet );
		}
		return sheetList;
	}

	/**
	 * 生成表头的样式
	 * 
	 * @param workbook
	 * @return HSSFFont
	 */
	public static HSSFCellStyle creatTitleStyle( HSSFWorkbook workbook ) {
		// 生成一个样式
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		style.setFillForegroundColor( HSSFColor.SKY_BLUE.index );
		style.setFillPattern( HSSFCellStyle.SOLID_FOREGROUND );
		style.setBorderBottom( HSSFCellStyle.BORDER_THIN );
		style.setBorderLeft( HSSFCellStyle.BORDER_THIN );
		style.setBorderRight( HSSFCellStyle.BORDER_THIN );
		style.setBorderTop( HSSFCellStyle.BORDER_THIN );
		style.setAlignment( HSSFCellStyle.ALIGN_CENTER );
		// 生成一个字体
		HSSFFont font = workbook.createFont();
		font.setColor( HSSFColor.VIOLET.index );
		font.setFontHeightInPoints( ( short )12 );
		font.setBoldweight( HSSFFont.BOLDWEIGHT_BOLD );
		// 把字体应用到当前的样式
		style.setFont( font );
		return style;
	}
	public static XSSFCellStyle creatTitleStyle( XSSFWorkbook workbook ) {
		// 生成一个样式
		XSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		style.setFillForegroundColor( HSSFColor.SKY_BLUE.index );
		style.setFillPattern( XSSFCellStyle.SOLID_FOREGROUND );
		style.setBorderBottom( XSSFCellStyle.BORDER_THIN );
		style.setBorderLeft( XSSFCellStyle.BORDER_THIN );
		style.setBorderRight( XSSFCellStyle.BORDER_THIN );
		style.setBorderTop( XSSFCellStyle.BORDER_THIN );
		style.setAlignment( XSSFCellStyle.ALIGN_CENTER );
		// 生成一个字体
		XSSFFont font = workbook.createFont();
		font.setColor( HSSFColor.VIOLET.index );
		font.setFontHeightInPoints( ( short )12 );
		font.setBoldweight( XSSFFont.BOLDWEIGHT_BOLD );
		// 把字体应用到当前的样式
		style.setFont( font );
		return style;
	}

	/**
	 * 生成表体的样式
	 * 
	 * @param workbook
	 * @return HSSFFont
	 */
	public static HSSFCellStyle creatDataStyle( HSSFWorkbook workbook ) {
		HSSFCellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor( HSSFColor.LIGHT_YELLOW.index );
		style.setFillPattern( HSSFCellStyle.SOLID_FOREGROUND );
		style.setBorderBottom( HSSFCellStyle.BORDER_THIN );
		style.setBorderLeft( HSSFCellStyle.BORDER_THIN );
		style.setBorderRight( HSSFCellStyle.BORDER_THIN );
		style.setBorderTop( HSSFCellStyle.BORDER_THIN );
		style.setAlignment( HSSFCellStyle.ALIGN_CENTER );
		style.setVerticalAlignment( HSSFCellStyle.VERTICAL_CENTER );
		style.setWrapText(true);//先设置为自动换行
		// 生成另一个字体
		HSSFFont font = workbook.createFont();
		font.setBoldweight( HSSFFont.BOLDWEIGHT_NORMAL );
		// 把字体应用到当前的样式
		style.setFont( font );
		// 设置单元格格式为文本(如果单元格内容是3306，在默认格式下excel文件会自动将3006修改成3306.0)
		HSSFDataFormat format = workbook.createDataFormat();
		style.setDataFormat( format.getFormat( "@" ) );
		return style;
	}

	public static XSSFCellStyle creatDataStyle( XSSFWorkbook workbook ) {
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor( HSSFColor.LIGHT_YELLOW.index );
		style.setFillPattern( XSSFCellStyle.SOLID_FOREGROUND );
		style.setBorderBottom( XSSFCellStyle.BORDER_THIN );
		style.setBorderLeft( XSSFCellStyle.BORDER_THIN );
		style.setBorderRight( XSSFCellStyle.BORDER_THIN );
		style.setBorderTop( XSSFCellStyle.BORDER_THIN );
		style.setAlignment( XSSFCellStyle.ALIGN_CENTER );
		style.setVerticalAlignment( XSSFCellStyle.VERTICAL_CENTER );
		// 生成另一个字体
		XSSFFont font = workbook.createFont();
		font.setBoldweight( XSSFFont.BOLDWEIGHT_NORMAL );
		// 把字体应用到当前的样式
		style.setFont( font );
		return style;
	}

	/**
	 * 生成头标题的样式，边框，加粗，字体20
	 *
	 * @param workbook
	 * @return HSSFFont
	 */

	public static HSSFCellStyle creatHeaderTitleStyle( HSSFWorkbook workbook ) {
		//生成样式对象
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);//设置上边框
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);   //设置下边框
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN); //设置做边框
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);    //设置右边框
		style.setWrapText(true);//先设置为自动换行
		// 生成另一个字体
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 15);//字体大小
		font.setFontName("Calibri");
		font.setColor(HSSFColor.BLACK.index);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把字体应用到当前的样式
		style.setFont( font );
		// 设置单元格格式为文本(如果单元格内容是3306，在默认格式下excel文件会自动将3006修改成3306.0)
		HSSFDataFormat format = workbook.createDataFormat();
		style.setDataFormat( format.getFormat( "@" ) );
		return style;
	}

	/**
	 * 生成头内容的样式，边框，字体12
	 *
	 * @param workbook
	 * @return HSSFFont
	 */
	public static HSSFCellStyle creatHeaderContentStyle( HSSFWorkbook workbook ) {
		//生成样式对象
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);//设置上边框
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);   //设置下边框
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN); //设置做边框
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);    //设置右边框
		style.setWrapText(true);//先设置为自动换行
		// 生成另一个字体
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);//字体大小
		font.setFontName("Calibri");
		font.setColor(HSSFColor.BLACK.index);
		// 把字体应用到当前的样式
		style.setFont( font );
		// 设置单元格格式为文本(如果单元格内容是3306，在默认格式下excel文件会自动将3006修改成3306.0)
		HSSFDataFormat format = workbook.createDataFormat();
		style.setDataFormat( format.getFormat( "@" ) );
		return style;
	}



	/**
	 * 对sheet页，写入数据
	 * 
	 * @param sheet
	 * @param titleArray
	 *            表头的数据
	 * @param titleStyle
	 * @param data
	 *            表体的数据
	 * @param dataStyle
	 */
	@SuppressWarnings( "deprecation" )
	public static void exportExcelData( HSSFWorkbook workbook, HSSFSheet sheet, String[] titleArray, HSSFCellStyle titleStyle, List<String[]> dataList, HSSFCellStyle dataStyle ) {

		// 产生表格标题行
		HSSFRow row = sheet.createRow( 0 );
		if( titleArray != null && titleArray.length != 0 ) {
			for( short i = 0; i < titleArray.length; i++ ) {
				HSSFCell cell = row.createCell( i );
				cell.setCellStyle( titleStyle );
				HSSFRichTextString text = new HSSFRichTextString( titleArray[ i ] );
				cell.setCellValue( text );
			}
		}

		// 遍历集合数据，产生数据行
		if( dataList.size() != 0 ) {
			for( int i = 0; i < dataList.size(); i++ ) {
				row = sheet.createRow( i + 1 );
				String[] dataArray = dataList.get( i );
				if( dataArray != null ) {
					for( int k = 0; k < dataArray.length; k++ ) {
						HSSFCell cell = row.createCell( ( short )k );
						cell.setCellStyle( dataStyle );
						cell.setCellValue( dataArray[ k ] );
					}
				}
			}
		}
	}
	public static void exportExcelData( XSSFWorkbook workbook, XSSFSheet sheet, String[] titleArray, XSSFCellStyle titleStyle, List<String[]> dataList, XSSFCellStyle dataStyle ) {

		// 产生表格标题行
		XSSFRow row = sheet.createRow( 0 );
		if( titleArray != null && titleArray.length != 0 ) {
			for( short i = 0; i < titleArray.length; i++ ) {
				XSSFCell cell = row.createCell( i );
				cell.setCellStyle( titleStyle );
				XSSFRichTextString text = new XSSFRichTextString( titleArray[ i ] );
				cell.setCellValue( text );
			}
		}

		// 遍历集合数据，产生数据行
		if( dataList.size() != 0 ) {
			for( int i = 0; i < dataList.size(); i++ ) {
				row = sheet.createRow( i + 1 );
				String[] dataArray = dataList.get( i );
				if( dataArray != null ) {
					for( int k = 0; k < dataArray.length; k++ ) {
						XSSFCell cell = row.createCell( ( short )k );
						cell.setCellStyle( dataStyle );
						cell.setCellValue( dataArray[ k ] );
					}
				}
			}
		}
	}

	public static void creatRowTextType( HSSFSheet sheet, Integer startrow, Integer endrow, Integer colnum ) {
		if( null != startrow && null != endrow && null != colnum && startrow < endrow ) {
			for( int i = startrow; i < endrow; i++ ) {
				HSSFRow createRow = sheet.createRow( startrow );
				for( int j = 0; j < colnum; j++ ) {
					HSSFCell createCell = createRow.createCell( j );
					createCell.setCellType( HSSFCell.CELL_TYPE_STRING );
				}
			}
		}
	}

	/**
	 * 拼接标题行
	 * 
	 * @param workbook
	 * @param sheet
	 * @param titleArray
	 * @param titleStyle
	 * @param startindex从哪个索引出进行拼接
	 */
	public static void exportExcelData( HSSFWorkbook workbook, HSSFSheet sheet, String[] titleArray, HSSFCellStyle titleStyle, Integer startindex ) {

		// 产生表格标题行
		HSSFRow row = sheet.getRow( 0 );
		Integer cellindex = startindex;
		if( titleArray != null && titleArray.length != 0 ) {
			for( short i = 0; i < titleArray.length; i++ ) {
				HSSFCell cell = row.createCell( cellindex );
				cell.setCellStyle( titleStyle );
				HSSFRichTextString text = new HSSFRichTextString( titleArray[ i ] );
				cell.setCellValue( text );
				cellindex++;
			}
		}
	}

	@SuppressWarnings( "deprecation" )
	public static void writeExcelData( HSSFWorkbook workbook, HSSFSheet sheet, HSSFCellStyle titleStyle, List<String[]> dataList, HSSFCellStyle dataStyle ) {
		// 遍历集合数据，产生数据行
		if( dataList.size() != 0 ) {
			for( int i = 0; i < dataList.size(); i++ ) {
				HSSFRow row = sheet.createRow( i );
				String[] dataArray = dataList.get( i );
				if( dataArray != null ) {
					for( int k = 0; k < dataArray.length; k++ ) {
						HSSFCell cell = row.createCell( ( short )k );
						cell.setCellStyle( dataStyle );
						cell.setCellValue( dataArray[ k ] );
					}
				}
			}
		}
	}

	/**
	 * 创建下拉菜单
	 * 
	 * @param sheet
	 * @param selectArray
	 *            下拉框中的内容
	 * @param ColumnNum
	 *            下拉框创建于第几列（从0开始）
	 * @param rowNum
	 *            下拉模型创建的个数
	 * @param selectedValue
	 *            下拉框中选中的内容
	 */
	public static void makeHSSFSelectedBox( HSSFSheet sheet, String[] selectArray, short ColumnNum, int rowNum, String selectedValue, HSSFCellStyle dataStyle ) {
		for( int i = 0; i < rowNum; i++ ) {
			@SuppressWarnings( "deprecation" )
			HSSFRow hssfrow = sheet.getRow( ( short )i + 1 );
			HSSFCell cell = hssfrow.getCell( ColumnNum );
			cell.setCellValue( selectedValue );
			CellRangeAddressList regions = new CellRangeAddressList( i + 1, i + 1, ColumnNum, ColumnNum );
			// 生成下拉框内容
			DVConstraint constraint = DVConstraint.createExplicitListConstraint( selectArray );
			// 绑定下拉框和作用区域
			HSSFDataValidation validation = new HSSFDataValidation( regions, constraint );
			// 对sheet页生效
			sheet.addValidationData( validation );
			cell.setCellValue( selectedValue );
			cell.setCellStyle( dataStyle );
		}
	}
	public static void makeHSSFSelectedBox( XSSFSheet sheet, String[] selectArray, short ColumnNum, int rowNum, String selectedValue, XSSFCellStyle dataStyle ) {
		for( int i = 0; i < rowNum; i++ ) {
			@SuppressWarnings( "deprecation" )
			XSSFRow hssfrow = sheet.getRow( ( short )i + 1 );
			XSSFCell cell = hssfrow.getCell( ColumnNum );
			cell.setCellValue( selectedValue );
			CellRangeAddressList addressList = new CellRangeAddressList( i + 1, i + 1, ColumnNum, ColumnNum );
			// 生成下拉框内容
			XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
			// 绑定下拉框和作用区域
			XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.createExplicitListConstraint(selectArray);
			XSSFDataValidation validation =(XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);
			// 对sheet页生效
			sheet.addValidationData( validation );
			cell.setCellValue( selectedValue );
			cell.setCellStyle( dataStyle );
		}
	}

	/**
	 * 创建下拉框，数据来源于另一个sheet
	 * 
	 * @param sheet
	 * @param sheetName
	 * @param ColumnNum
	 * @param rowNum
	 * @param selectedValue
	 * @param dataStyle
	 */
	public static void makeHSSFSelectedBoxFromOtherSheet( HSSFSheet sheet, String sheetName, short ColumnNum, int rowNum, String selectedValue, HSSFCellStyle dataStyle ) {
		for( int i = 0; i < rowNum; i++ ) {
			@SuppressWarnings( "deprecation" )
			HSSFRow hssfrow = sheet.getRow( ( short )i + 1 );
			HSSFCell cell = hssfrow.getCell( ColumnNum );
			cell.setCellValue( selectedValue );
			CellRangeAddressList regions = new CellRangeAddressList( i + 1, i + 1, ColumnNum, ColumnNum );
			// 生成下拉框内容
			DVConstraint constraint = DVConstraint.createFormulaListConstraint( sheetName );
			// 绑定下拉框和作用区域
			HSSFDataValidation validation = new HSSFDataValidation( regions, constraint );
			// 对sheet页生效
			sheet.addValidationData( validation );
			cell.setCellValue( selectedValue );
			cell.setCellStyle( dataStyle );
		}
	}

	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param textlist
	 *            下拉框显示的内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFValidation( HSSFSheet sheet, String[] textlist, int firstRow, int endRow, int firstCol, int endCol ) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createExplicitListConstraint( textlist );
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList( firstRow, endRow, firstCol, endCol );
		// 数据有效性对象
		HSSFDataValidation data_validation_list = new HSSFDataValidation( regions, constraint );
		sheet.addValidationData( data_validation_list );
		return sheet;
	}

	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param textlist
	 *            下拉框显示的内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFValidation( HSSFSheet sheet, String[] textlist, String promptTitle, String promptContent, int firstRow, int endRow, int firstCol, int endCol ) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createExplicitListConstraint( textlist );
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList( firstRow, endRow, firstCol, endCol );
		// 数据有效性对象
		HSSFDataValidation data_validation_list = new HSSFDataValidation( regions, constraint );
		data_validation_list.createPromptBox( promptTitle, promptContent );
		sheet.addValidationData( data_validation_list );
		return sheet;
	}

	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param textlist
	 *            下拉框显示的内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFValidation( HSSFSheet sheet, String sheetName, int firstRow, int endRow, int firstCol, int endCol ) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createFormulaListConstraint( sheetName );
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList( firstRow, endRow, firstCol, endCol );
		// 数据有效性对象
		HSSFDataValidation data_validation_list = new HSSFDataValidation( regions, constraint );
		sheet.addValidationData( data_validation_list );
		return sheet;
	}

	/**
	 * 设置单元格上提示
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param promptTitle
	 *            标题
	 * @param promptContent
	 *            内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFPrompt( HSSFSheet sheet, String promptTitle, String promptContent, int firstRow, int endRow, int firstCol, int endCol ) {
		// 构造constraint对象
		DVConstraint constraint = DVConstraint.createCustomFormulaConstraint( "BB1" );
		// 四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList( firstRow, endRow, firstCol, endCol );
		// 数据有效性对象
		HSSFDataValidation data_validation_view = new HSSFDataValidation( regions, constraint );
		data_validation_view.createPromptBox( promptTitle, promptContent );
		sheet.addValidationData( data_validation_view );
		return sheet;
	}

	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param textlist
	 *            下拉框显示的内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFValidation( HSSFSheet sheet, String sheetName, String promptTitle, String promptContent, int firstRow, int endRow, int firstCol, int endCol ) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createFormulaListConstraint( sheetName );
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList( firstRow, endRow, firstCol, endCol );
		// 数据有效性对象
		HSSFDataValidation data_validation_list = new HSSFDataValidation( regions, constraint );
		data_validation_list.createPromptBox( promptTitle, promptContent );
		sheet.addValidationData( data_validation_list );
		return sheet;
	}

}
