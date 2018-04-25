package online.yaosiyaoqi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by alfredo on 2016/10/24.
 */
@Retention( value = RetentionPolicy.RUNTIME )
@Target( ElementType.FIELD)
public @interface ExcelExportValidate {


	int firstRow();
	int endRow();
	int firstCol();
	int endCol();
	boolean droplists() default false;
}
