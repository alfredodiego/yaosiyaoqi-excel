package online.yaosiyaoqi.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
  * sheet页签
  * @author 	alfredo
  * @date   	2016/8/17 16:04
  */
@Retention(value = RetentionPolicy.RUNTIME)
@Target( ElementType.TYPE)
public @interface TableExportSheet {
	/**
	 * sheet名称
	 * @return
	 */
	String name();

	/**
	 * 是否隐藏sheet
	 * 默认可见
	 * @return
	 */
	boolean visible() default true;
	/**
	 * 头实现
	 * 默认可见
	 * @return
	 */
	Class Header() default TableHeaderInterface.class;



}
