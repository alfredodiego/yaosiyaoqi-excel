package online.yaosiyaoqi.excel;

import java.util.List;

/**
 *
 *
 * 导入导出的解析参数类
 * @Title: suninfo
 * @Package com.suninfo.util.excel
 * Company:suninfo
 * @author alfredo
 * @date 2016/8/22
 */
public class ExportParseParam {

	private List<? extends Object> data;
	private Class clazz;
	private Object  headerData;

	public List<? extends Object> getData() {
		return data;
	}

	public void setData( List<? extends Object> data ) {
		this.data = data;
	}

	public Class getClazz() {
		return clazz;
	}

	public void setClazz( Class clazz ) {
		this.clazz = clazz;
	}

	public Object getHeaderData() {
		return headerData;
	}

	public void setHeaderData(Object headerData) {
		this.headerData = headerData;
	}
}
