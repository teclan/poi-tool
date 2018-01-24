package poi.lx.utils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

/**
 * 对一些常见值的处理
 * 
 * @author teclan
 * 
 *         email: tbj621@163.com
 *
 *         2017年10月27日
 */
public class Objects {
	@SuppressWarnings("unused")
	private static final Logger LOGGER = LoggerFactory.getLogger(Objects.class);

	public static boolean isNull(Object value) {
		return value == null;
	}

	public static boolean isNotNull(Object value) {
		return !isNull(value);
	}

	public static boolean isNotNullString(Object value) {
		return !isNullString(value);
	}

	/**
	 * 字符串是否是 null 或者 ""
	 * 
	 * @param value
	 * @return
	 */
	public static boolean isNullString(Object value) {
		return (value == null || "".equals(value.toString().trim()));
	}

	public static JSONObject removeNullKey(JSONObject jsonObject) {

		List<String> deleted = new ArrayList<String>();

		for (String key : jsonObject.keySet()) {

			if (isNull(jsonObject.get(key))) {
				deleted.add(key);
			}
		}
		for (String key : deleted) {
			jsonObject.remove(key);
		}

		return jsonObject;
	}

	public static JSONObject setNull2EmptyString(JSONObject jsonObject) {

		List<String> keys = new ArrayList<String>();

		for (String key : jsonObject.keySet()) {

			if (isNull(jsonObject.get(key))) {
				keys.add(key);
			}
		}
		for (String key : keys) {
			jsonObject.put(key, "");
		}

		return jsonObject;
	}

	public static boolean hasNullValue(JSONObject jsonObject) {
		for (String key : jsonObject.keySet()) {
			if (isNullString(jsonObject.get(key))) {
				return true;
			}
		}
		return false;
	}

	public static List<String> getKeys(JSONObject jsonObject) {

		List<String> keys = new ArrayList<String>();

		for (String key : jsonObject.keySet()) {
			keys.add(key);
		}
		return keys;
	}

	public static String Joiner(String separator, List<String> collection) {

		if (collection.isEmpty() || collection.size() == 0) {
			return "";
		}

		StringBuffer sb = new StringBuffer();

		if (collection.size() == 1) {
			return collection.iterator().next();
		} else {
			Iterator<String> iterator = collection.iterator();
			while (iterator.hasNext()) {
				sb.append(iterator.next()).append(separator);
			}
		}
		String result = sb.toString();

		return result.substring(0, result.length() - separator.length());
	}

	public static String Joiner(String separator, Set<String> collection) {

		if (collection.isEmpty() || collection.size() == 0) {
			return "";
		}

		StringBuffer sb = new StringBuffer();

		if (collection.size() == 1) {
			return collection.iterator().next();
		} else {
			Iterator<String> iterator = collection.iterator();
			while (iterator.hasNext()) {
				sb.append(iterator.next()).append(separator);
			}
		}
		String result = sb.toString();

		return result.substring(0, result.length() - separator.length());
	}

	public static String JoinerForDeleteSql(String logic, Map<String, Object> map) {

		StringBuffer sb = new StringBuffer();

		for (String key : map.keySet()) {

			if (map.get(key) == null) {
				continue;
			}

			sb.append(String.format(" %s = '%s' %s", key, map.get(key).toString(), logic));
		}

		String result = sb.toString();

		if (result.indexOf(logic) < 0) {
			return " 1=2 ";

		}

		return result.substring(0, result.lastIndexOf(logic));
	}

	public static String JoinerForSql(String logic, String column, Set<String> collection) {

		if (collection.isEmpty() || collection.size() == 0) {
			return " 1 = 1";
		}

		StringBuffer sb = new StringBuffer();

		if (collection.size() == 1) {
			return sb.append(String.format("( %s = '%s' )", column, collection.iterator().next())).toString();
		} else {
			sb.append(String.format(" %s = ", column));

			Iterator<String> iterator = collection.iterator();
			while (iterator.hasNext()) {
				sb.append("'").append(iterator.next()).append("'").append(String.format(" %s %s = ", logic, column));
			}
		}
		String result = sb.toString();

		return " ( " + result.substring(0, result.lastIndexOf(logic)) + " ) ";
	}

	public static String JoinerForSql(String logic, String column, JSONArray collection) {

		if (collection.isEmpty() || collection.size() == 0) {
			return " 1 = 1";
		}

		StringBuffer sb = new StringBuffer();

		if (collection.size() == 1) {
			return sb.append(String.format("( %s = '%s' )", column, collection.iterator().next())).toString();
		} else {
			sb.append(String.format(" %s = ", column));

			for (int i = 0; i < collection.size(); i++) {
				sb.append("'").append(collection.get(i).toString()).append("'")
						.append(String.format(" %s %s = ", logic, column));
			}
		}
		String result = sb.toString();

		return " ( " + result.substring(0, result.lastIndexOf(" " + logic + " ")) + " ) ";
	}

	public static boolean isNull(List<?> list) {
		return list == null || list.isEmpty();
	}

	public static boolean isNotNull(List<?> list) {
		return !isNull(list);
	}

	public static Map<String, Object> removeUnnecessaryColumns(List<String> columns,
			Map<String, Object> namesAndValues) {
		List<String> delete = new ArrayList<String>();

		for (String key : namesAndValues.keySet()) {
			if (!columns.contains(key)) {
				delete.add(key);
			}
		}

		for (String key : delete) {
			namesAndValues.remove(key);
		}

		return namesAndValues;

	}

	public static boolean isNumeric(String str) {

		if (str == null || "".equals(str.trim())) {
			return false;
		}

		Pattern pattern = Pattern.compile("[0-9]*");
		Matcher isNum = pattern.matcher(str);
		if (!isNum.matches()) {
			return false;
		}
		return true;
	}
}
