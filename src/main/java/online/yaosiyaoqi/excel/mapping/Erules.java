package online.yaosiyaoqi.excel.mapping;

import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;
import java.util.List;

/**
 * 导入导出excel xml
 * @Title: Erules
 * @Package com.suninfo.util.excel.mapping
 * Company:suninfo
 * @author alfredo
 * @date 2016/3/30 16:54
 */
@XmlRootElement(name = "erules")
@XmlAccessorType(value = javax.xml.bind.annotation.XmlAccessType.NONE)
public class Erules implements Serializable{
	private static final long serialVersionUID = 1L;
	/**
	 * 对应的mapping
	 */
	private List<Erule> eruleList;
	@XmlElement(name = "erule")
	public List<Erule> getEruleList() {
		return eruleList;
	}
	public void setEruleList(List<Erule> eruleList) {
		this.eruleList = eruleList;
	}
	

}
