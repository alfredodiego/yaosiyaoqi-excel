package online.yaosiyaoqi.excel.annotation;



import online.yaosiyaoqi.excel.Locator;
import online.yaosiyaoqi.excel.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static java.text.MessageFormat.format;

/**
 *
 * excel注解通用的工具方法
 * @Title: suninfo
 * @Package com.suninfo.util.excel.annotations
 * Company:suninfo
 * @author alfredo
 * @date 2016/8/24
 */
public class ExcelAnnotationUtil {
	private static DataFormatter formatter = new DataFormatter();

	/**
	 * 获取excel单元格field，并且根据columnorder进行排序，由小到大
	 * @param clazz
	 * @return
	 */
	public static List<Field> getExcelExportCellFieldList(Class clazz){
		List<Field> excelColumns = new ArrayList<>();
		List<SortField> sortFields = new ArrayList<>();
		for (Field field : clazz.getDeclaredFields()) {
			TableExportCell cellAnnotation = field.getAnnotation(TableExportCell.class);
			if (cellAnnotation != null) {
				int order = cellAnnotation.columnOrder();
				field.setAccessible(true);
				SortField sf = new SortField();
				sf.order = order;
				sf.field = field;
				sortFields.add(sf);
			}
		}
		Collections.sort( sortFields );
		for( SortField sortField : sortFields ) {
			excelColumns.add( sortField.field );
		}
		return excelColumns;
	}

	public static List<Field> getExcelExcelExportValidateFieldList(Class clazz){
		List<Field> excelColumns = new ArrayList<>();
		for (Field field : clazz.getDeclaredFields()) {
			ExcelExportValidate cellAnnotation = field.getAnnotation(ExcelExportValidate.class);
			if (cellAnnotation != null) {
				field.setAccessible( true );
				excelColumns.add( field );
			}
		}
		return excelColumns;
	}
	private static class SortField implements Comparable{
		Integer order;
		Field field;
		@Override
		public int compareTo( Object o ) {
			if(o instanceof SortField){
				SortField sf = (SortField) o;
				return order.compareTo( sf.order );
			}
			return 0;
		}
	}

	public static <T> T getCellValue(Sheet sheet, Class<T> type, Integer row, Integer col, boolean zeroIfNull) {
		Cell cell = getCell(sheet, row, col);

		if (type.equals(String.class)) {
			return (T) getStringCell(cell);
		}

		if (type.equals(Date.class)) {
			return cell == null ? null : (T) getDateCell(cell, new Locator(sheet.getSheetName(), row, col));
		}

		if (type.equals(Integer.class)) {
			return (T) getIntegerCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col));
		}

		if (type.equals(Double.class)) {
			return (T) getDoubleCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col));
		}

		if (type.equals(Long.class)) {
			return (T) getLongCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col));
		}

		if (type.equals(BigDecimal.class)) {
			return (T) getBigDecimalCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col));
		}

		return null;
	}
	private static BigDecimal getBigDecimalCell(Cell cell, boolean zeroIfNull, Locator locator ) {
		String val = getStringCell(cell);
		if(val == null || val.trim().equals("")) {
			if(zeroIfNull) {
				return BigDecimal.ZERO;
			}
			return null;
		}
		try {
			return new BigDecimal(val);
		} catch (NumberFormatException e) {
		}

		if (zeroIfNull) {
			return BigDecimal.ZERO;
		}
		return null;
	}

	public static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
		Row row = sheet.getRow(rowNumber - 1);
		return row == null ? null : row.getCell(columnNumber - 1);
	}

	public static String getStringCell(Cell cell) {
		if (cell == null) {
			return null;
		}

		if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			int type = cell.getCachedFormulaResultType();

			if (type == HSSFCell.CELL_TYPE_NUMERIC) {
				FormulaEvaluator fe = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
				return formatter.formatCellValue(cell, fe);
			}

			if (type == HSSFCell.CELL_TYPE_ERROR) {
				return "";
			}

			if (type == HSSFCell.CELL_TYPE_STRING) {
				return cell.getRichStringCellValue().getString().trim();
			}

			if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
				return "" + cell.getBooleanCellValue();
			}

		} else if (cell.getCellType() != HSSFCell.CELL_TYPE_NUMERIC) {
			return cell.getRichStringCellValue().getString().trim();
		}

		return formatter.formatCellValue(cell);
	}

	public static Date getDateCell(Cell cell, Locator locator) {
		try {
			if (!HSSFDateUtil.isCellDateFormatted(cell)) {
				throwExcelParsingException( locator );
			}
			return HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
		} catch (IllegalStateException illegalStateException) {
			throwExcelParsingException( locator );
		}
		return null;
	}

	public static Double getDoubleCell(Cell cell, boolean zeroIfNull, Locator locator) {
		if (cell == null) {
			return zeroIfNull ? 0d : null;
		}

		if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC || cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			return cell.getNumericCellValue();
		}

		if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
			return zeroIfNull ? 0d : null;
		}

		return null;
	}

	public static Long getLongCell(Cell cell, boolean zeroIfNull, Locator locator) {
		Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator);
		return doubleValue == null ? null : doubleValue.longValue();
	}

	public static Integer getIntegerCell(Cell cell, boolean zeroIfNull, Locator locator) {
		Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator);
		return doubleValue == null ? null : doubleValue.intValue();
	}

	private static Double getNumberWithoutDecimals(Cell cell, boolean zeroIfNull, Locator locator) {
		Double doubleValue = getDoubleCell(cell, zeroIfNull, locator);
		if (doubleValue != null && doubleValue % 1 != 0) {
		}
		return doubleValue;
	}

	private static void throwExcelParsingException(Locator locator){
		throw new ExcelParsingException(format("无效的数据在sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
	}
}
