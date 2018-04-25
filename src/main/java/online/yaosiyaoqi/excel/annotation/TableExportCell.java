package online.yaosiyaoqi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Table单元格
 * @author alfredo
 * @date 2016/8/17 16:03
 */
@Retention( value = RetentionPolicy.RUNTIME )
@Target( ElementType.FIELD )
public @interface TableExportCell {

	/**
	 * 列顺序
	 * @return
	 */
	int columnOrder();

	/**
	 * 列名称
	 * @return
	 */
	String headerName();

	/**
	 * 动态表头
	 * @return
	 */
	boolean dynamic() default false;

}
