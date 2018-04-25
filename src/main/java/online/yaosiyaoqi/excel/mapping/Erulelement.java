package online.yaosiyaoqi.excel.mapping;

import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;

/**
 * 导入导出excel xml
 * @Title: Erulelement
 * @Package com.suninfo.util.excel.mapping
 * Company:suninfo
 * @author alfredo
 * @date 2016/3/30 16:54
 */
@XmlRootElement(name = "erulelement")
@XmlAccessorType(value = javax.xml.bind.annotation.XmlAccessType.NONE)
public class Erulelement implements Serializable{
	private static final long serialVersionUID = 1L;
	/**
	 * excel列索引，从0开始，中间无间隔
	 */
	@XmlAttribute
	private Integer index;
	/**
	 * 实体类中的属性名
	 */
	@XmlAttribute
	private String column;
	/**
	 * excel中的标题名
	 */
	@XmlAttribute
	private String value;
	public Integer getIndex() {
		return index;
	}
	public void setIndex(Integer index) {
		this.index = index;
	}
	public String getColumn() {
		return column;
	}
	public void setColumn(String column) {
		this.column = column;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	

}
