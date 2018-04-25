package online.yaosiyaoqi.excel.parser;

import online.yaosiyaoqi.excel.ExportParseParam;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 *
 * @Title: suninfo
 * @Package com.suninfo.util.excel.parser
 * Company:suninfo
 * @author alfredo
 * @date 2016/8/24
 */
public interface ExcelExportParseInter {

	/**
	 *  excel导出解析，会生成文件
	  * @param		data sheet数据
	  * @param		clazz 解析的类别
	  * @param
	  * @return
	  * @throws		
	  * @author 	alfredo
	  * @date   	2016/8/24 13:31
	  */
	Workbook parse(List<? extends Object> data, Class clazz, Object headData) ;

	/**
	 * 多个sheet调用此接口
	 * @param parseParamlist
	 * @return
	 */
	Workbook parse(List<ExportParseParam> parseParamlist) ;

}
