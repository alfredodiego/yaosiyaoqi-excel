package online.yaosiyaoqi.excel.exception;

/**
 *
 * @Title: suninfo
 * @Package com.suninfo.util.excel
 * Company:suninfo
 * @author alfredo
 * @date 2016/8/23
 */
public class ExcelParseException extends Exception{

	public ExcelParseException() {
		super();
	}

	public ExcelParseException( String message ) {
		super( message );
	}

	public ExcelParseException( String message, Throwable cause ) {
		super( message, cause );
	}

	public ExcelParseException( Throwable cause ) {
		super( cause );
	}

	protected ExcelParseException( String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace ) {
		super( message, cause, enableSuppression, writableStackTrace );
	}
}
