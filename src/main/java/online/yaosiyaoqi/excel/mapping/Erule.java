package online.yaosiyaoqi.excel.mapping;

import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;
import java.util.List;

/**
 * 导入导出excel xml
 * @Title: Erule
 * @Package com.suninfo.util.excel.mapping
 * Company:suninfo
 * @author alfredo
 * @date 2016/3/30 16:54
 */
@XmlRootElement(name = "erule")
@XmlAccessorType(value = javax.xml.bind.annotation.XmlAccessType.NONE)
public class Erule implements Serializable{
	private static final long serialVersionUID = 1L;
	/**
	 * 唯一表示一个模型
	 */
	@XmlAttribute
	private String key;
	/**
	 * xml对应的实体类的全路径
	 */
	@XmlAttribute
	private String type;
	/**
	 * 实体中的元素
	 */
	@XmlElement(name = "erulelement")
	private List<Erulelement> erulelementList;
	@XmlElement(name = "other")
	private Other other;

	public List<Erulelement> getErulelementList() {
		return erulelementList;
	}

	public void setErulelementList(List<Erulelement> erulelementList) {
		this.erulelementList = erulelementList;
	}

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public Other getOther() {
		return other;
	}

	public void setOther(Other other) {
		this.other = other;
	}
	

}
