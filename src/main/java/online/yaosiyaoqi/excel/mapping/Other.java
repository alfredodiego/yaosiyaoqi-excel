package online.yaosiyaoqi.excel.mapping;

import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;

/**
 * 封装那些没有字段的导入数据
 * @Title: Other
 * @Package com.suninfo.util.excel.mapping
 * Company:suninfo
 * @author alfredo
 * @date 2016/3/30 16:54
 */
@XmlRootElement(name = "other")
@XmlAccessorType(value = javax.xml.bind.annotation.XmlAccessType.NONE)
public class Other  implements Serializable{
	private static final long serialVersionUID = 1L;
	@XmlAttribute
	private String column;

	public String getColumn() {
		return column;
	}

	public void setColumn(String column) {
		this.column = column;
	}
	
}
