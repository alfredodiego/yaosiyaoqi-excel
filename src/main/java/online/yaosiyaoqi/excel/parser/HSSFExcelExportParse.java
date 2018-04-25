package online.yaosiyaoqi.excel.parser;


import online.yaosiyaoqi.excel.ExcelUtil;
import online.yaosiyaoqi.excel.ExportParseParam;
import online.yaosiyaoqi.excel.annotation.*;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.reflect.Field;
import java.util.*;

/**
 *
 * 2003excel文档导出解析类
 * @Title: suninfo
 * @Package com.suninfo.util.excel.parser
 * Company:suninfo
 * @author alfredo
 * @date 2016/8/24
 */
public class HSSFExcelExportParse implements ExcelExportParseInter {
	private final Log log = LogFactory.getLog( HSSFExcelExportParse.class );
	private HSSFWorkbook workbook;
	@Override
	public Workbook parse( List<? extends Object> data, Class clazz,Object headData)  {
		if (data == null || data.isEmpty()) {
			return null;
		}
		workbook = new HSSFWorkbook();
		//解析sheet数据
		parseData( data,clazz, headData);
		return workbook;
	}
	@Override
	public Workbook parse( List<ExportParseParam> parseParamlist )  {
		if( CollectionUtils.isNotEmpty( parseParamlist )){
			workbook = new HSSFWorkbook();
			for( ExportParseParam param : parseParamlist ) {
				parseData( param.getData(),param.getClazz(),param.getHeaderData());
			}
			return workbook;
		}
		return null;
	}

	private void parseData(List<? extends Object> data, Class clazz,Object headData)  {
		HSSFSheet sheet = createHSSFSheet( clazz, workbook );
		//解析非头部信息
		TableExportSheet sheetAnnotation = (TableExportSheet ) clazz.getAnnotation(TableExportSheet.class);
		List<Field> excelColumns = ExcelAnnotationUtil.getExcelExportCellFieldList( clazz );
		Class header = sheetAnnotation.Header();
		Integer rowCount = 0;
		TableHeaderInterface instant = createInstant(header);
		if(instant != null){
			rowCount =  instant.creatExcelHeader(headData, sheet, workbook,excelColumns.size());
		}
		//行数，第一行是title，之后的数据
		rowCount = createHSSFTitle( ExcelUtil.creatTitleStyle( workbook ), sheet, excelColumns, rowCount,data.get( 0 ));
		createHSSFData( data, ExcelUtil.creatDataStyle( workbook ), sheet, excelColumns, rowCount );
		List<Field> validateFieldList = ExcelAnnotationUtil.getExcelExcelExportValidateFieldList( clazz );
		createHSSFValidate(sheet,data.get( 0 ),validateFieldList);
	}

	private TableHeaderInterface createInstant(Class header) {
		try {
			TableHeaderInterface o = (TableHeaderInterface )header.newInstance();
			return o;

		} catch (InstantiationException e) {
			log.error("生成实体错误",e);
		} catch (IllegalAccessException e) {
			log.error("生成实体错误",e);
		}
		return null;

	}

	private void createHSSFValidate( Sheet sheet,Object obj,List<Field> validateFieldList ) {
		//关联下拉列表数据名称
		String nameName = "hidden";
		//隐藏索引
		int hIndex=0;
		for( Field field : validateFieldList ) {
			ExcelExportValidate validate = field.getAnnotation( ExcelExportValidate.class );
			if(null!=validate && validate.droplists() && null!=obj){
				try {
					Object oj = field.get( obj );
					if(oj instanceof List){
						List<String> dataoption = ( List<String> )oj;
						if(CollectionUtils.isNotEmpty( dataoption )){
							hIndex+=1;
							dataValidations( sheet,dataoption,validate.firstRow(),validate.endRow(),validate.firstCol(),validate.endCol() ,nameName+hIndex,hIndex);
						}
					}
				} catch( IllegalAccessException e ) {
					e.printStackTrace();
				}

			}
		}

	}

	private HSSFSheet createHSSFSheet( Class clazz, HSSFWorkbook workbook ) {
		TableExportSheet sheetAnnotation = (TableExportSheet ) clazz.getAnnotation(TableExportSheet.class);
		HSSFSheet sheet;
		if (sheetAnnotation == null){
			sheet = workbook.createSheet();
		}else{
			sheet = workbook.createSheet(sheetAnnotation.name());
			boolean visible = sheetAnnotation.visible();
			if(!visible){
				workbook.setSheetHidden( workbook.getSheetIndex( sheetAnnotation.name() ),true );
			}
		}
		return sheet;
	}

	private int createHSSFTitle( CellStyle headerStyle, HSSFSheet sheet, List<Field> excelColumns, int rowCount ,Object data) {
		int colCount = 0;
		HSSFRow headerRow = sheet.createRow(rowCount++);
		for (Field header : excelColumns) {
			TableExportCell cellAnnotation = header.getAnnotation(TableExportCell.class);
			if(cellAnnotation.dynamic()){
				try {
					Object map = header.get( data );
					if(map instanceof LinkedHashMap){
						LinkedHashMap<String,String> head = ( LinkedHashMap<String, String> )map;
						Set<String> strings = head.keySet();
						for( String string : strings ) {
							HSSFCell cell = headerRow.createCell(colCount++);
							cell.setCellValue(string);
							cell.setCellStyle(headerStyle);
						}
					}
				} catch( IllegalAccessException e ) {
					log.error( "反射获取注解对应值时参数无法被访问,动态列头问题",e );
					e.printStackTrace();
				}

			}else{
				String headerLabel = cellAnnotation.headerName();
				HSSFCell cell = headerRow.createCell(colCount++);
				cell.setCellValue(headerLabel);
				cell.setCellStyle(headerStyle);
			}

		}
		return rowCount;
	}

	private void createHSSFData(List<? extends Object> data, HSSFCellStyle dataStyle, HSSFSheet sheet, List<Field> excelColumns, int rowCount ) {
		try {
			int colCount;
			Map<String,Integer> macMap=new HashMap<>();
			for (int i = 0; i <excelColumns.size() ; i++) {
				macMap.put("max"+i,20);
			}
			for (Object dataObject : data) {
				HSSFRow dataRow = sheet.createRow(rowCount++);
				colCount = 0;
				for (int i = 0; i <excelColumns.size() ; i++) {
					Integer maxValue = macMap.get("max" + i);
					Field field =excelColumns.get(i);
					Object obj = field.get( dataObject );
					TableExportCell cellAnnotation = field.getAnnotation(TableExportCell.class);
					if(cellAnnotation.dynamic()){
						if(obj instanceof LinkedHashMap){
							LinkedHashMap<String,String> head = ( LinkedHashMap<String, String> )obj;
							Collection<String> values = head.values();
							int dynaCnt = excelColumns.size();
							for( String value : values ) {
								HSSFCell cell = dataRow.createCell(colCount++);
								cell.setCellValue(new HSSFRichTextString(value));
								cell.setCellStyle(dataStyle);
								sheet.setColumnWidth(dynaCnt,  (int)((maxValue + 0.72) * 256));
								dynaCnt++;
							}
						}
					}else {
						String value = "";
						if(obj != null){
							value = String.valueOf( obj );
						}
						HSSFCell cell = dataRow.createCell(colCount++);
						cell.setCellValue(new HSSFRichTextString(value));
						if (maxValue<value.length()){
							maxValue=value.length();
							macMap.put("max"+i,maxValue);
						}
						cell.setCellStyle(dataStyle);
					}
				}
			}
			for (int i = 0; i <excelColumns.size() ; i++) {
				Integer maxValue = macMap.get("max" + i);
				sheet.setColumnWidth(i,  (int)((maxValue + 0.72) * 256));
			}
		} catch( IllegalAccessException e ) {
			log.error( "反射获取注解对应值时参数无法被访问",e );
			e.printStackTrace();
		}

	}

	private void dataValidations(Sheet sheet,List<String> options,int firstRow, int endRow, int firstCol, int endCol ,String nameName,int hindex){
		if(CollectionUtils.isNotEmpty( options )){
			HSSFSheet hidden = workbook.createSheet(nameName);
			for (int i = 0, length= options.size(); i < length; i++) {
				String name = options.get( i );
				HSSFRow row = hidden.createRow(i);
				HSSFCell cell = row.createCell(0);
				cell.setCellValue(name);
			}
			Name namedCell = workbook.createName();
			namedCell.setNameName(nameName);
			namedCell.setRefersToFormula(nameName+"!$A$1:$A$" + options.size());
			DVConstraint constraint = DVConstraint.createFormulaListConstraint(nameName);
			CellRangeAddressList addressList = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
			HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
			workbook.setSheetHidden(hindex, true);
			sheet.addValidationData(validation);
		}




	}



}
