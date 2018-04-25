package online.yaosiyaoqi.excel.mapping;




import online.yaosiyaoqi.excel.XMLUtil;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * 用于导入导出时对于对应的配置规则应用工具
 * @Title: ExcelMappingUtil
 * @Package com.suninfo.util.excel.mapping
 * Company:suninfo
 * @author alfredo
 * @date 2016/3/30 16:54
 */
public class ExcelMappingUtil {
	private static Erules eruls;

	private static final String DEFAULT_ERULS_CFG = "/config/excel-mappingrules.xml";

	/**
	 * 所有excel导入规则列表
	 */
	private static List<Erule> eruleList;

	/**
	 * 封装了KEY对应的excel规则
	 */
	private static Map<String, Erule> erulemap;

	/**
	 * 封装了key对应的规则中的字段名
	 */
	private static Map<String, List<Map<Integer, String>>> columsMap;

	/**
	 * 封装了key对应的规则中表头名称
	 */
	private static Map<String, List<Map<Integer, String>>> valuesMap;

	static {
		/**
		 * 初始化加载xml中的规则
		 */
		if (eruls == null) {
			InputStream ins = ExcelMappingUtil.class.getResourceAsStream(DEFAULT_ERULS_CFG);
			eruls = XMLUtil.formXmlToObject(ins, Erules.class);
			if (null != eruls && null != eruls.getEruleList()) {
				eruleList = eruls.getEruleList();
				erulemap = new HashMap<String, Erule>();
				columsMap = new HashMap<String, List<Map<Integer, String>>>();
				valuesMap = new HashMap<String, List<Map<Integer, String>>>();
				for (Erule e : eruleList) {
					String key = e.getKey();
					if ( null!=key) {
						erulemap.put(e.getKey(), e);
					}
					if (null != e.getErulelementList()) {
						List<Erulelement> erulelementList = e.getErulelementList();
						List<Map<Integer, String>> columlist = new ArrayList<Map<Integer, String>>();
						List<Map<Integer, String>> valueslist = new ArrayList<Map<Integer, String>>();
						Map<Integer, String> columMap = new HashMap<Integer, String>();
						Map<Integer, String> valueMap = new HashMap<Integer, String>();
						for (Erulelement el : erulelementList) {
							columMap.put(el.getIndex(), el.getColumn());
							valueMap.put(el.getIndex(), el.getValue());
							columlist.add(columMap);
							valueslist.add(valueMap);
						}
						if (null!=key) {
							columsMap.put(key, columlist);
							valuesMap.put(key, valueslist);
						}

					}
				}
			}
		}

	}

	/**
	  * 【描述】：	 获取所有的规则列表
	  * 【步骤】：
	  * @param
	  * @return
	  * @throws
	  * @author 	alfredo
	  * @date   	2016/4/5 14:22
	  */
	public static Erules getErulesErules() {
		return eruls;
	}


	/**
	 * 【描述】：	获取最大索引
	 * 【步骤】：
	 * @param
	 * @return
	 * @throws
	 * @author 	alfredo
	 * @date   	2016/4/5 14:22
	 */
	public static Integer getMaxIndex(String key){
		Erule erule = erulemap.get(key);
		if(null!=erule){
			return erule.getErulelementList().size();
		}else{
			return null;
		}
	}


	/**
	 * 【描述】：	根据配置key获取对应的实体对象
	 * 【步骤】：
	 * @param
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws ClassNotFoundException
	 * @author 	alfredo
	 * @date   	2016/4/5 14:22
	 */
	public static Object createIntanceByruleType(String key)
			throws InstantiationException, IllegalAccessException, ClassNotFoundException {
		if (null != erulemap && null != erulemap.get(key)) {
			Erule e = erulemap.get(key);
			if (null!=e.getType()) {
				return Class.forName(e.getType()).newInstance();
			}
		}
		return null;
	}


	/**
	  * 【描述】：	根据key获取定义的表头
	  * 【步骤】：
	  * @param
	  * @return
	  * @throws
	  * @author 	alfredo
	  * @date   	2016/4/5 14:30
	  */
	public static String[] getTableHeadBykey(String key) {
		List<String> head = new ArrayList<String>();
		if (null != valuesMap && null != valuesMap.get(key)) {
			List<Map<Integer, String>> list = valuesMap.get(key);
			if (null != list && list.size() > 0) {
				Map<Integer, String> map = list.get(0);
				Set<Integer> keySet = map.keySet();
				for (int i = 0; i <  keySet.size();i++) {
					head.add(map.get(i));
				}
			}
		}
		String[] headarray = new String[head.size()];
		for (int i = 0; i < head.size(); i++) {
			headarray[i] = head.get(i);
		}
		return headarray;
	}

	/**
	 * 【描述】：	获取对应的字段名称
	 * 【步骤】：
	 * @param
	 * @return
	 * @throws
	 * @author 	alfredo
	 * @date   	2016/4/5 14:30
	 */
	public static String getColumn(String key, Integer index) {
		if (null != columsMap && null!=key) {
			List<Map<Integer, String>> list = columsMap.get(key);
			if (null != list && list.size() > 0) {
				for (Map<Integer, String> map : list) {
					return map.get(index);
				}
			}

		}
		return null;
	}

	public static Map<Integer, String> getColumnMap(String key) {
		if (null != columsMap && null!=key) {
			List<Map<Integer, String>> list = columsMap.get(key);
			if (null != list && list.size() > 0) {
				return list.get(0);
			}
		}
		return null;
	}


	/**
	 * 【描述】：	根据配置的excelmapping模型去反射返回需要的数据,调用此方法，需要业务层进行异常处理
	 * 【步骤】：
	 * @param
	 * @return
	 * @throws
	 * @author 	alfredo
	 * @date   	2016/4/5 14:30
	 */
	public static List<String[]> getCellDatalist(String key, List<?> objlist) throws Exception{
		List<String[]> datalist = new ArrayList<String[]>();
		if (null != erulemap && null != erulemap.get(key)) {
			Erule e = erulemap.get(key);
			if (null!=e.getType()) {
				Class<?> classType = Class.forName(e.getType());
				Map<Integer, String> columnMap = ExcelMappingUtil.getColumnMap(key);
				Set<Integer> keys = columnMap.keySet();
				for(Object obj : objlist){
					String[] celldata = new String[keys.size()];
					for (Integer index : keys) {
						Method declaredMethod = classType.getDeclaredMethod("get"+ capitalize(columnMap.get(index)));
						Object result = declaredMethod.invoke(obj);
						if(null!=result){
							if(result instanceof String){
								celldata[index] = (String) result;
							}
						}
					}
					Erule erule = erulemap.get(key);
					Other other = erule.getOther();
					if(null!=other && null!=other.getColumn()){
						String fieldName = other.getColumn();
						Method declaredMethod = classType.getDeclaredMethod("get"+capitalize(fieldName));
						Object result = declaredMethod.invoke(obj);
						String [] re = null;
						if(null!=result){
							if(result instanceof LinkedList){
								LinkedList li = (LinkedList)result;
								if(null!=li){
									re = new String [li.size()];
									for(int i=0;i<li.size();i++){
										re[i] = (String)li.get(i);
									}
								}
							}
						}
						if(null!=re){
							String[] joinedArray = ( String[] )Array.newInstance(celldata.getClass().getComponentType(), celldata.length + re.length);
							System.arraycopy(celldata, 0, joinedArray, 0, celldata.length);
							System.arraycopy(re, 0, joinedArray, celldata.length, re.length);
							datalist.add(joinedArray);
						}else{
							datalist.add(celldata);
						}
					}else{
						datalist.add(celldata);
					}
				}

			}

		}
		return datalist;
	}
	public static String capitalize(String str) {
		int strLen;
		return str != null && (strLen = str.length()) != 0 ? (new StringBuffer(strLen)).append(Character.toTitleCase(str.charAt(0))).append(str.substring(1)).toString() : str;
	}

	/**
	 * 【描述】：	excel中的数据转换成对象
	 * 【步骤】：
	 * @param
	 * @return
	 * @throws
	 * @author 	alfredo
	 * @date   	2016/4/5 14:30
	 */
	public static List<Object> getCellModellist(String key, List<LinkedList<String>> dataList) throws Exception{
		List<Object> objlist = new ArrayList<Object>();
		Map<Integer, String> columnMap = getColumnMap(key);
		Set<Integer> keyset = columnMap.keySet();
		if(null!=dataList && dataList.size()>0){
			for(int i=1;i<dataList.size();i++){
				LinkedList<String> cellarray = dataList.get(i);
				Object createIntanceByruleType = createIntanceByruleType(key);
				Class<? extends Object> clazz = createIntanceByruleType.getClass();
				for(int j=0;j<cellarray.size();j++){
					if(keyset.contains(j)){
						String val = cellarray.get(j);
						String fieldname = columnMap.get(j);
						Field field = clazz.getDeclaredField(fieldname);
						Method declaredMethod = clazz.getDeclaredMethod("set"+capitalize(fieldname),field.getType());
						declaredMethod.invoke(createIntanceByruleType, val);
					}

				}
				packageOther(key, createIntanceByruleType, cellarray);
				objlist.add(createIntanceByruleType);
			}
		}
		return objlist;
	}

	private static void packageOther(String key,Object instance,LinkedList<String> list) throws Exception{
		Erule erule = erulemap.get(key);
		Other other = erule.getOther();
		if(null!=other && null!=other.getColumn()){
			Class<? extends Object> clazz = instance.getClass();
			String fieldName = other.getColumn();
			Field field = clazz.getDeclaredField(fieldName);
			Method declaredMethod = clazz.getDeclaredMethod("set"+capitalize(fieldName),field.getType());
			declaredMethod.invoke(instance, list);
		}
	}
}
