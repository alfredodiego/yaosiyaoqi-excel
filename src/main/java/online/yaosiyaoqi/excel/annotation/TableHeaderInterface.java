package online.yaosiyaoqi.excel.annotation;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Created by huwl on 2017/4/12.
 */
public interface TableHeaderInterface {

	public Integer creatExcelHeader(Object data, HSSFSheet sheet, HSSFWorkbook wb, int size);

	default Integer defaultHeader() {return 0;}
}
