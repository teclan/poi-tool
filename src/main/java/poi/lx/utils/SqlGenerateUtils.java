package poi.lx.utils;

import java.util.Arrays;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.google.common.base.Joiner;

/**
 * 一个通过给定 json 或 json 的 map<String,Object> 表示，得到对应的插入，更新以及条件查询SQL和对应的值
 * 
 * @author teclan
 * 
 *         email: tbj621@163.com
 *
 *         2017年10月10日
 */
public class SqlGenerateUtils {
	private static final Logger LOGGER = LoggerFactory.getLogger(SqlGenerateUtils.class);

	// 区间查询使用 intervals 的json数组表示
	// { "intervals":[{"column":"birthday","begin":"1999","end":"2003" }]}
	private static final String INTERVALS = "intervals";
	private static final String COLUMN = "column";
	private static final String BEGIN = "begin";
	private static final String END = "end";

	/**
	 * 获取区间查询的条件值，默认模糊查询, 与 getIntervalQuerySql(String json) 搭配使用
	 * 
	 * @return
	 */
	public static String[] getIntervalQueryValues(String json) {
		return getIntervalValues(json, false);

	}

	/**
	 * 获取区间查询的条件值
	 * 
	 * @param exact
	 *            是否精确查询，模糊查询 exact 设为 false;
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static String[] getIntervalValues(String json, boolean exact) {

		Map<String, Object> map = (Map<String, Object>) JSON.parse(json);

		JSONArray intervals = JSON.parseArray(map.remove(INTERVALS).toString());

		String[] values = new String[intervals.size() * 2];

		for (int i = 0, j = 0; i < intervals.size(); i++) {
			values[j] = exact ? ((JSONObject) intervals.get(i)).getString(BEGIN)
					: "%" + ((JSONObject) intervals.get(i)).getString(BEGIN) + "%";

			values[j + 1] = exact ? ((JSONObject) intervals.get(i)).getString(END)
					: "%" + ((JSONObject) intervals.get(i)).getString(END) + "%";
		}

		return values;
	}

	/**
	 * 获取查询的条件值 ， 默认模糊查询，与 getSimpleQuerySql(String json) 搭配使用
	 * 
	 * @return
	 */
	public static Object[] getSimpleQueryValues(String json) {
		return getSimpleValues(json, false);

	}

	/**
	 * 获取非区间查询的条件值，与 getSimpleQuerySql(String json,boolean exact) 搭配使用
	 * 
	 * @param exact
	 *            是否精确查询，模糊查询 exact 设为 false;
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static Object[] getSimpleValues(String json, boolean exact) {

		Map<String, Object> map = (Map<String, Object>) JSON.parse(json);

		map.remove(INTERVALS);

		Object[] values = new Object[map.size()];

		int index = 0;
		for (String key : map.keySet()) {
			values[index] = exact ? map.get(key).toString() : "%" + map.get(key).toString() + "%";
			index++;
		}
		return values;
	}

	/**
	 * 获取非区间查询条件的SQL,默认模糊查询
	 * 
	 * @param json
	 * @return
	 */
	public static String getSimpleQuerySql(String json) {
		return getSimpleQuerySql(json, false);
	}

	/**
	 * 获取非区间查询条件的SQL
	 * 
	 * @param json
	 * @param exact
	 *            是否精确查询
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static String getSimpleQuerySql(String json, boolean exact) {

		Map<String, Object> map = (Map<String, Object>) JSON.parse(json);

		map.remove(INTERVALS);

		String[] columns = new String[map.size()];

		int index = 0;
		for (String key : map.keySet()) {
			columns[index] = exact ? String.format("%s = ?", key) : String.format("%s like ?", key);
			index++;
		}

		return Joiner.on(" or ").join(columns);
	}

	/**
	 * 获取区间查询条件的SQL
	 * 
	 * @param json
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static String getIntervalQuerySql(String json) {

		Map<String, Object> map = (Map<String, Object>) JSON.parse(json);

		JSONArray intervals = JSON.parseArray(map.remove(INTERVALS).toString());

		String[] columns = new String[intervals.size()];

		for (int i = 0; i < intervals.size(); i++) {
			columns[i] = String.format(" %s betwee ? and ?", ((JSONObject) intervals.get(i)).getString(COLUMN));
		}

		return Joiner.on(" and ").join(columns);
	}

	@SuppressWarnings("unchecked")
	public static Object[] getInsertValues(String json) {

		Map<String, Object> namesAndValues = (Map<String, Object>) JSON.parse(json);

		Object[] values = new Object[namesAndValues.size()];
		int index = 0;

		for (String key : namesAndValues.keySet()) {
			values[index] = namesAndValues.get(key);
			index++;
		}

		return values;
	}

	/**
	 * 获取插入的所有值，与 generateSqlForInsert(Map<String, Object> namesAndValues) 搭配使用
	 * 
	 * @param namesAndValues
	 * @return
	 */
	public static Object[] getInsertValues(Map<String, Object> namesAndValues) {
		Object[] values = new Object[namesAndValues.size()];
		int index = 0;

		for (String key : namesAndValues.keySet()) {
			values[index] = namesAndValues.get(key);
			index++;
		}

		return values;
	}

	/**
	 * 获取插入的SQL
	 * 
	 * @param namesAndValues
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static String generateSqlForInsert(String json) {

		Map<String, Object> namesAndValues = (Map<String, Object>) JSON.parse(json);

		String[] columns = new String[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			columns[index] = key;
			index++;
		}

		return String.format(" (%s) values (%s) ", Joiner.on(",").join(columns), getFillMark(columns.length));
	}

	/**
	 * 获取插入的SQL
	 * 
	 * @param namesAndValues
	 * @return
	 */
	public static String generateSqlForInsert(Map<String, Object> namesAndValues) {

		String[] columns = new String[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			columns[index] = key;
			index++;
		}

		return String.format(" (%s) values (%s) ", Joiner.on(",").join(columns), getFillMark(columns.length));
	}

	private static String getFillMark(int count) {

		String[] marks = new String[count];
		Arrays.fill(marks, "?");

		return Joiner.on(",").join(marks);
	}

	/**
	 * 获取更新的SQL
	 * 
	 * @param nameAndValues
	 * @return
	 */
	public static String generateSqlForUpdate(Map<String, Object> namesAndValues) {

		String[] columns = new String[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			columns[index] = key;
			index++;
		}

		return Joiner.on(" = ? ,").join(columns) + " = ? ";
	}

	@SuppressWarnings("unchecked")
	public static String generateSqlForUpdate(String json) {

		Map<String, Object> namesAndValues = (Map<String, Object>) JSON.parse(json);

		String[] columns = new String[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			columns[index] = key;
			index++;
		}
		return Joiner.on(" = ? ,").join(columns) + " = ? ";

	}

	/**
	 * 获取被更新字段的值，与 generateSqlForUpdate(Map<String, Object> nameAndValues) 搭配使用;
	 * 
	 * @param namesAndValues
	 * @return
	 */
	public static Object[] getNewValuesForUpdate(Map<String, Object> namesAndValues) {

		Object[] values = new Object[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			values[index] = namesAndValues.get(key).toString();
			index++;
		}
		return values;
	}

	@SuppressWarnings("unchecked")
	public static Object[] getNewValuesForUpdate(String json) {

		Map<String, Object> namesAndValues = (Map<String, Object>) JSON.parse(json);

		Object[] values = new Object[namesAndValues.size()];

		int index = 0;

		for (String key : namesAndValues.keySet()) {
			values[index] = namesAndValues.get(key).toString();
			index++;
		}
		return values;
	}

	public static Object[] getNewValuesForUpdate(Object... namesAndValues) {

		Object[] values = new Object[namesAndValues.length / 2];

		for (int i = 0, j = 1; i < values.length; i++, j += 2) {
			values[i] = namesAndValues[j];
		}
		return values;
	}

	public static String generateSqlForUpdate(Object... nameAndValues) {

		String[] columns = new String[nameAndValues.length / 2];

		for (int i = 0, j = 0; i < columns.length; i++, j += 2) {
			columns[i] = (String) nameAndValues[j];
		}
		return Joiner.on(" = ? ,").join(columns) + " = ? ";
	}

}
