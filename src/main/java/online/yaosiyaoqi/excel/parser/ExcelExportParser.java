package online.yaosiyaoqi.excel.parser;


import online.yaosiyaoqi.excel.ExcelUtil;
import online.yaosiyaoqi.excel.ExportParseParam;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelExportParser{

	private final Log log = LogFactory.getLog( ExcelExportParser.class );
	private ExcelExportParseInter parseInter;
	private File file;

	public ExcelExportParser(){
		parseInter = new HSSFExcelExportParse();
	}

	public ExcelExportParser(String fileName) throws IOException {
		if( null!=fileName && fileName.length()>0){
			int lastIndexOf = fileName.lastIndexOf( "." );
			if(-1 == lastIndexOf ){
				throw new IllegalArgumentException( "文件后缀名不支持");
			}
			String filesubfix = fileName.substring( lastIndexOf +1, fileName.length() );
			if( ExcelUtil.SUFFIXES2003.equalsIgnoreCase( filesubfix )){
				parseInter = new HSSFExcelExportParse();
			}else if(ExcelUtil.SUFFIXES2007.equalsIgnoreCase( filesubfix )){
				parseInter = new XSSFExcelExportParse();
			}else{
				log.error("创建ExcelExportParser对象参数非法：文件后缀名不支持");
				throw new IllegalArgumentException( "文件后缀名不支持");
			}
		}
		file = new File(fileName);
		if (!file.exists()) {
			file.createNewFile();
		}
	}

	/**
	 * 解析多个sheet
	 * @param parseParamlist
	 */
	public boolean parseSheet(List<ExportParseParam>  parseParamlist){
		boolean flag = false;
		Workbook workbook = parseInter.parse( parseParamlist);
		if(null!=workbook){
			try (FileOutputStream fileOutputStream = new FileOutputStream(file)){
				workbook.write(fileOutputStream);
				flag = true;
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return flag;
	}

	/**
	 * 解析单个sheet
	 * @param data
	 * @param t
	 */
	public boolean parseSheet(List<? extends Object> data, Class t,Object headData) {
		Workbook workbook = parseInter.parse( data,t,headData );
		try (FileOutputStream fileOutputStream = new FileOutputStream(file)){
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return true;
	}


	public byte[] getParsedWorkbook(List<? extends Object> data, Class t,Object headData) {
		Workbook parse = parseInter.parse( data, t,headData );
		try (ByteArrayOutputStream os = new ByteArrayOutputStream();){
			parse.write(os);
			byte[] bytes = os.toByteArray();
			return bytes;
		} catch( IOException e ) {
			log.error( "excel导出数据，io流write异常",e );
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 获取多个sheet
	 * @param parseParamlist
	 */
	public byte[] getParsedManyWorkbook(List<ExportParseParam>  parseParamlist){
		Workbook workbook = parseInter.parse( parseParamlist);
		try (ByteArrayOutputStream os = new ByteArrayOutputStream();){
			workbook.write(os);
			byte[] bytes = os.toByteArray();
			return bytes;
		} catch( IOException e ) {
			log.error( "excel导出数据，io流write异常",e );
			e.printStackTrace();
		}
		return null;
	}

}
