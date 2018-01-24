package poi.lx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.typesafe.config.Config;
import com.typesafe.config.ConfigFactory;

import poi.lx.db.DatabaseServer;

public class ExcelUtils {
	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);
	private static SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

	private static final String MANUFACTURER = "P2P_NPC";
	private static final String DEVMODELID = "900000001";
	private static final String IMM_DEV_INFO = "imm_devinfo";
	private static final String IMM_CAMERA = "imm_camera";
	private static final String IMM_NETNVR_ATTR = "imm_netnvrattr";
	private static final String INSERT_SQL = "insert into %s %s";

	/*
	 * 000004用户，有4个通道：
	 * 
	 * NVR：13010200001180000001
	 * 
	 * 通道1：13010200001310000001
	 * 
	 * 通道2：13010200001310000002
	 * 
	 * 通道3：13010200001310000003
	 * 
	 * 通道4：13010200001310000004
	 * 
	 * 000005用户，有4个通道：
	 * 
	 * NVR：13010200001180000002
	 * 
	 * 通道1：13010200001310000005
	 * 
	 * 通道2：13010200001310000006
	 * 
	 * 通道3：13010200001310000007
	 * 
	 * 通道4：13010200001310000008
	 */

	// 2018 年1月12号迁移的NVR数据，以下编号记得保留
	private static int START_CAMEAR_ID = 201801120;
	private static int START_NVR_GBID = 80000001;
	private static int START_CAMERA_GBID = 10000001;

	private static String DEV_LNG;
	private static String DEV_LAT;
	private static String AREAID;
	private static String VIDEO_SERVER;
	private static String PLATFORM_ID;

	static {
		File file = new File("config/application.conf");
		Config root = ConfigFactory.parseFile(file);
		Config config = root.getConfig("config");

		VIDEO_SERVER = config.getString("videoServer");
		AREAID = config.getString("areaId");
		DEV_LNG = config.getString("lng");
		DEV_LAT = config.getString("lat");
		PLATFORM_ID = config.getString("platformId");

	}

	private static String getNextCameraGbId() {
		return "130102000013" + (START_CAMERA_GBID++);
	}

	private static String getNextNVRGbId() {
		return "130102000011" + (START_NVR_GBID++);
	}

	private static int getNextCameraId() {
		return START_CAMEAR_ID++;
	}

	public static void analyzeNOArch(String filePath) {

		Workbook wb = null;
		InputStream input = null;
		try {
			File file = new File(filePath);
			input = new FileInputStream(file);// 读取文件流
		} catch (IOException e1) {
			LOGGER.error(e1.getMessage(), e1);
		}
		try {
			wb = WorkbookFactory.create(input); // 构建excel文件
		} catch (Exception e1) {
			LOGGER.error(e1.getMessage(), e1);
		}

		for (int sheetNum = 0; sheetNum < 2; sheetNum++) {
			// for (int sheetNum = 0; sheetNum < 1; sheetNum++) {
			Sheet sheet = wb.getSheetAt(sheetNum);
			LOGGER.info("sheet {},name:{}", sheetNum, sheet.getSheetName());

			int LastRow = sheet.getLastRowNum();

			for (int i = 1; i <= LastRow; i++) {
				Row row = sheet.getRow(i);

				// 设备编号
				String devId = getPhoneValueString(row.getCell(0));

				devId = devId.replace("“", "").replace("‘", "").replace("’", "").replace("", "").replace(";", "");

				if (Objects.isNullString(devId)) {
					continue;
				}

				if (sheetNum == 1) {
					devId = getDevIdForPingShan(devId);
				} else {
					devId = getDevId(devId);
				}

				if (devId.contains("_error") || devId.contains("退网")) {
					LOGGER.info("设备编号无效，dev：{}", getPhoneValueString(row.getCell(0)));
					continue;
				}

				// 设备名称
				String devName = getValue(row.getCell(1));

				// 视频通道
				String videoChannel = getValue(row.getCell(2));
				// devTUTKID
				String devTUTKID = getValue(row.getCell(4));

				if (Objects.isNullString(devTUTKID) || devTUTKID.indexOf("无") >= 0 || devTUTKID.indexOf("——") >= 0) {
					LOGGER.info("设备 {} 云ID无效，云ID:{}", devId, devTUTKID);
					continue;
				}

				// 设备登录用户（未知）
				String devLoginName = "admin";
				// 设备登录密码
				String devLoginPwd = getPassword(getPhoneValueString(row.getCell(5)));
				// 国标ID
				String gbId = getNextNVRGbId();

				Map<String, Object> devData = new HashMap<String, Object>();
				devData.put("devId", devId);
				devData.put("devName", devName);
				devData.put("define5", gbId);
				devData.put("pnlActID", devId);
				devData.put("areaId", AREAID);
				devData.put("devType", 10);
				// 设备型号，使用字典中唯一的型号
				devData.put("devModelId", DEVMODELID);
				devData.put("instMan", "");
				devData.put("devLng", DEV_LNG);
				devData.put("devlat", DEV_LAT);
				devData.put("fMemo", "");
				devData.put("manufacturer", MANUFACTURER);
				devData.put("platformId", PLATFORM_ID);

				devData.put("devInstDate", "2018-01-01");
				devData.put("pnlAddr", "");
				devData.put("instUnit", "");

				LOGGER.info("\n========= NVR ==========");

				try {
					// TODO
					// 插入设备基本信息表
					DatabaseServer.getDb().exec(
							String.format(INSERT_SQL, IMM_DEV_INFO, SqlGenerateUtils.generateSqlForInsert(devData)),
							SqlGenerateUtils.getInsertValues(devData));
					LOGGER.info("插入设备基本信息表成功，devId:{}...", devId);
				} catch (Exception e) {
					LOGGER.error(e.getMessage(), e);
				}

				Map<String, Object> netnvrattrData = new HashMap<String, Object>();
				netnvrattrData.put("devId", devId);
				netnvrattrData.put("devLoginName", devLoginName);
				netnvrattrData.put("devLoginPwd", devLoginPwd);
				netnvrattrData.put("devTUTKID", devTUTKID);
				netnvrattrData.put("videoServer", VIDEO_SERVER);

				try {
					// TODO
					// 插入互联网属性表
					DatabaseServer.getDb()
							.exec(String.format(INSERT_SQL, IMM_NETNVR_ATTR,
									SqlGenerateUtils.generateSqlForInsert(netnvrattrData)),
									SqlGenerateUtils.getInsertValues(netnvrattrData));

					LOGGER.info("插入互联网属性表成功，devId:{}...", devId);
				} catch (Exception e) {
					LOGGER.error(e.getMessage(), e);
				}

				videoChannel = videoChannel.indexOf(".") >= 0 ? videoChannel.substring(0, videoChannel.indexOf("."))
						: videoChannel;

				// 构造监控点
				if (Objects.isNullString(videoChannel) || !Objects.isNumeric(videoChannel)) {
					LOGGER.info("设备 {} 无通道...", devId);
					continue;
				}

				LOGGER.info("\n========= 摄像机==========");
				for (int channelId = 0; channelId < Double.valueOf(videoChannel); channelId++) {
					Map<String, Object> cameraDevData = new HashMap<String, Object>();
					String cameraDevId = getNextCameraId() + "";
					cameraDevData.put("devId", cameraDevId);
					cameraDevData.put("devName", devName + "_" + channelId);
					cameraDevData.put("define5", getNextCameraGbId());
					cameraDevData.put("pnlActID", devId);
					cameraDevData.put("areaId", AREAID);
					cameraDevData.put("devType", "12");
					cameraDevData.put("devModelId", DEVMODELID);
					cameraDevData.put("instMan", "");
					cameraDevData.put("devLng", "1.0");
					cameraDevData.put("devlat", "2.0");
					cameraDevData.put("relateNVR", devId);
					cameraDevData.put("instUnit", "");
					cameraDevData.put("fMemo", "");
					cameraDevData.put("manufacturer", MANUFACTURER);
					cameraDevData.put("platformId", PLATFORM_ID);

					try {
						// TODO
						// 插入设备信息表

						DatabaseServer.getDb()
								.exec(String.format(INSERT_SQL, IMM_DEV_INFO,
										SqlGenerateUtils.generateSqlForInsert(cameraDevData)),
										SqlGenerateUtils.getInsertValues(cameraDevData));
						LOGGER.info("插入设备基本信息表成功，devId:{}...", cameraDevId);
					} catch (Exception e) {
						LOGGER.error(e.getMessage(), e);
					}

					Map<String, Object> cameraAttrData = new HashMap<String, Object>();

					cameraAttrData.put("devId", cameraDevId);
					cameraAttrData.put("devChannelId", channelId);
					cameraAttrData.put("cameraName", devName + "_" + channelId);
					cameraAttrData.put("relateNVR", devId);
					cameraAttrData.put("atPos", "未知");
					cameraAttrData.put("instDate", "");
					cameraAttrData.put("wantDo", "-1");
					cameraAttrData.put("almType", "-1");
					cameraAttrData.put("cameraModeId", "-1");
					cameraAttrData.put("cameraType", "-1");
					cameraAttrData.put("fMemo", "");
					cameraAttrData.put("devMonitorId", "000" + channelId);

					// String videoUrlSufString = ":9000/" + devTUTKID + ":0:" + MANUFACTURER + ":"
					// + channelId + ":0:"
					// + devLoginName + ":" + devLoginPwd + "/av_stream";
					String videoUrlSufString = ":9000/" + devTUTKID + ":0:" + MANUFACTURER + ":" + channelId + ":1:"
							+ devLoginName + ":" + devLoginPwd + "/av_stream";

					cameraAttrData.put("videoUrlSuf", videoUrlSufString);
					cameraAttrData.put("videoServer", VIDEO_SERVER);

					try {
						// TODO
						// 插入摄像机属性表
						DatabaseServer.getDb()
								.exec(String.format(INSERT_SQL, IMM_CAMERA,
										SqlGenerateUtils.generateSqlForInsert(cameraAttrData)),
										SqlGenerateUtils.getInsertValues(cameraAttrData));
						LOGGER.info("插入摄像机信息表成功，devId:{}...", cameraDevId);
					} catch (Exception e) {
						LOGGER.error(e.getMessage(), e);
					}
				}

			}
			LOGGER.info("\n=====================\n");
		}

	}

	public static void analyze(String filePath) {

		Workbook wb = null;
		InputStream input = null;
		try {
			File file = new File(filePath);
			input = new FileInputStream(file);// 读取文件流
		} catch (IOException e1) {
			LOGGER.error(e1.getMessage(), e1);
		}
		try {
			wb = WorkbookFactory.create(input); // 构建excel文件
		} catch (Exception e1) {
			LOGGER.error(e1.getMessage(), e1);
		}

		for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
			// for (int sheetNum = 0; sheetNum < 1; sheetNum++) {
			Sheet sheet = wb.getSheetAt(sheetNum);
			LOGGER.info("sheet {},name:{}", sheetNum, sheet.getSheetName());

			int LastRow = sheet.getLastRowNum();

			for (int i = 1; i <= LastRow; i++) {
				Row row = sheet.getRow(i);

				Calendar calendar = Calendar.getInstance();
				try {
					calendar.setTime(DATE_FORMAT.parse("1900-1-1"));
				} catch (ParseException e) {
					LOGGER.info(e.getMessage(), e);
				}

				if (row == null) {
					continue;
				}

				// 设备安装时间
				String devInstDate = "";

				// 设备编号
				String devId = getPhoneValueString(row.getCell(1));

				if (Objects.isNullString(devId)) {
					continue;
				}

				devId = devId.replace("“", "").replace("‘", "").replace("’", "").replace("", "").replace(";", "");

				// 用户编号
				String userId = devId = getDevId(devId);

				if (devId.contains("_error") || devId.contains("退网")) {
					LOGGER.info("设备编号无效，dev：{}", getPhoneValueString(row.getCell(1)));
					continue;
				}

				// 设备编号
				String devName = getValue(row.getCell(2));

				String cell1 = getValue(row.getCell(0));

				if (cell1.indexOf(".") > 0) {
					cell1 = cell1.substring(0, cell1.lastIndexOf("."));
				}

				if (Objects.isNumeric(cell1)) {
					// 日期需要 -2，不然结果正确，原因未知
					calendar.add(5, Integer.valueOf(cell1) - 2);
					devInstDate = DATE_FORMAT.format(calendar.getTime());
				} else {
					devInstDate = getValue(row.getCell(0)).replace(".", "-");
				}

				if (Objects.isNullString(devInstDate)) {
					continue;
				}

				// 用户名称
				String userName = getValue(row.getCell(2));
				// 用户地址
				String userAddr = getValue(row.getCell(3));
				// 负责人
				String contact = getValue(row.getCell(4));
				// 负责人电话
				String cPhone = getPhoneValueString(row.getCell(5));
				// 口令
				String payNO = getValue(row.getCell(6));
				// 联系人1
				String cName1 = getValue(row.getCell(7));
				// 联系人1电话
				String cphone1 = getPhoneValueString(row.getCell(8));
				// 联系人2
				String cName2 = getValue(row.getCell(9));
				// 联系人2电话
				String cphone2 = getPhoneValueString(row.getCell(10));
				// 主机型号
				String devModelName = getValue(row.getCell(11));
				// 设备位置
				String pnlAddr = getValue(row.getCell(12));
				// 模块编号（未知）
				String modelId = getValue(row.getCell(13));
				// 视频通道
				String videoChannel = getValue(row.getCell(14));
				// 防区资料
				String zoneDoc = getValue(row.getCell(15));
				// 设备自带的编号
				String devSn = getValue(row.getCell(16));
				// devTUTKID
				String devTUTKID = getValue(row.getCell(17));

				if (devTUTKID.indexOf(".") > 0) {
					devTUTKID = devTUTKID.substring(0, devTUTKID.indexOf("."));
				}

				if (Objects.isNullString(devTUTKID) || "0".equals(devTUTKID) || devTUTKID.indexOf("无") >= 0
						|| devTUTKID.indexOf("——") >= 0) {

					LOGGER.info("设备 {} 云ID无效，云ID:{}", devId, devTUTKID);
					continue;
				}

				// 设备登录用户（未知）
				String devLoginName = "admin";
				// 设备登录密码
				String devLoginPwd = getPassword(getPhoneValueString(row.getCell(18)));
				// 国标ID
				String gbId = getNextNVRGbId();

				Map<String, Object> devData = new HashMap<String, Object>();
				devData.put("devId", devId);
				devData.put("devName", devName);
				devData.put("define5", gbId);
				devData.put("pnlActID", devId);
				devData.put("areaId", AREAID);
				devData.put("devType", 10);
				// 设备型号，使用字典中唯一的型号
				devData.put("devModelId", DEVMODELID);
				devData.put("instMan", "");
				devData.put("devInstDate", devInstDate);
				devData.put("devLng", DEV_LNG);
				devData.put("devlat", DEV_LAT);
				devData.put("pnlAddr", pnlAddr);
				devData.put("instUnit", "");
				devData.put("fMemo", "");
				devData.put("manufacturer", MANUFACTURER);
				devData.put("platformId", PLATFORM_ID);

				LOGGER.info("\n========= NVR ==========");

				try {

					// TODO
					// 插入设备基本信息表
					DatabaseServer.getDb().exec(
							String.format(INSERT_SQL, IMM_DEV_INFO, SqlGenerateUtils.generateSqlForInsert(devData)),
							SqlGenerateUtils.getInsertValues(devData));
					LOGGER.info("插入设备基本信息表成功，devId:{}...", devId);
				} catch (Exception e) {
					LOGGER.error(e.getMessage(), e);
				}

				Map<String, Object> netnvrattrData = new HashMap<String, Object>();
				netnvrattrData.put("devId", devId);
				netnvrattrData.put("devLoginName", devLoginName);
				netnvrattrData.put("devLoginPwd", devLoginPwd);
				netnvrattrData.put("devTUTKID", devTUTKID);
				netnvrattrData.put("videoServer", VIDEO_SERVER);

				try {
					// TODO
					// 插入互联网属性表
					DatabaseServer.getDb()
							.exec(String.format(INSERT_SQL, IMM_NETNVR_ATTR,
									SqlGenerateUtils.generateSqlForInsert(netnvrattrData)),
									SqlGenerateUtils.getInsertValues(netnvrattrData));

					LOGGER.info("插入互联网属性表成功，devId:{}...", devId);
				} catch (Exception e) {
					LOGGER.error(e.getMessage(), e);
				}

				videoChannel = videoChannel.indexOf(".") >= 0 ? videoChannel.substring(0, videoChannel.indexOf("."))
						: videoChannel;

				// 构造监控点
				if (Objects.isNullString(videoChannel) || !Objects.isNumeric(videoChannel)) {
					LOGGER.info("设备 {} 无通道...", devId);
					continue;
				}

				LOGGER.info("\n========= 摄像机==========");
				for (int channelId = 0; channelId < Double.valueOf(videoChannel); channelId++) {
					Map<String, Object> cameraDevData = new HashMap<String, Object>();
					String cameraDevId = getNextCameraId() + "";
					cameraDevData.put("devId", cameraDevId);
					cameraDevData.put("devName", devName + "_" + channelId);
					cameraDevData.put("define5", getNextCameraGbId());
					cameraDevData.put("pnlActID", devId);
					cameraDevData.put("areaId", AREAID);
					cameraDevData.put("devType", "12");
					cameraDevData.put("devModelId", DEVMODELID);
					cameraDevData.put("instMan", "");
					cameraDevData.put("devInstDate", devInstDate);
					cameraDevData.put("devLng", "1.0");
					cameraDevData.put("devlat", "2.0");
					cameraDevData.put("relateNVR", devId);
					cameraDevData.put("pnlAddr", pnlAddr);
					cameraDevData.put("instUnit", "");
					cameraDevData.put("fMemo", "");
					cameraDevData.put("manufacturer", MANUFACTURER);
					cameraDevData.put("platformId", PLATFORM_ID);

					try {
						// TODO
						// 插入设备信息表

						DatabaseServer.getDb()
								.exec(String.format(INSERT_SQL, IMM_DEV_INFO,
										SqlGenerateUtils.generateSqlForInsert(cameraDevData)),
										SqlGenerateUtils.getInsertValues(cameraDevData));
						LOGGER.info("插入设备基本信息表成功，devId:{}...", cameraDevId);
					} catch (Exception e) {
						LOGGER.error(e.getMessage(), e);
					}

					Map<String, Object> cameraAttrData = new HashMap<String, Object>();

					cameraAttrData.put("devId", cameraDevId);
					cameraAttrData.put("devChannelId", channelId);
					cameraAttrData.put("cameraName", devName + "_" + channelId);
					cameraAttrData.put("relateNVR", devId);
					cameraAttrData.put("atPos", "未知");
					cameraAttrData.put("instDate", "");
					cameraAttrData.put("wantDo", "-1");
					cameraAttrData.put("almType", "-1");
					cameraAttrData.put("cameraModeId", "-1");
					cameraAttrData.put("cameraType", "-1");
					cameraAttrData.put("fMemo", "");
					cameraAttrData.put("devMonitorId", "000" + channelId);

					// String videoUrlSufString = ":9000/" + devTUTKID + ":0:" + MANUFACTURER + ":"
					// + channelId + ":0:"
					// + devLoginName + ":" + devLoginPwd + "/av_stream";
					String videoUrlSufString = ":9000/" + devTUTKID + ":0:" + MANUFACTURER + ":" + channelId + ":1:"
							+ devLoginName + ":" + devLoginPwd + "/av_stream";
					cameraAttrData.put("videoUrlSuf", videoUrlSufString);
					cameraAttrData.put("videoServer", VIDEO_SERVER);

					try {

						// TODO
						// 插入摄像机属性表
						DatabaseServer.getDb()
								.exec(String.format(INSERT_SQL, IMM_CAMERA,
										SqlGenerateUtils.generateSqlForInsert(cameraAttrData)),
										SqlGenerateUtils.getInsertValues(cameraAttrData));
						LOGGER.info("插入摄像机信息表成功，devId:{}...", cameraDevId);
					} catch (Exception e) {
						LOGGER.error(e.getMessage(), e);
					}
				}

			}
			LOGGER.info("\n=====================\n");

		}

	}

	private static String getDevIdForPingShan(String value) {

		if (value.length() == 6) {
			String prefix = "80003";

			return prefix + value.substring(2, value.length());
		} else {
			LOGGER.error("不是成都理想设备编号： {}", value);
			return value + "_error";
		}
	}

	private static String getDevId(String value) {

		if (value.length() == 6) {
			String prefix = "80002";

			return prefix + value.substring(2, value.length());
		} else {
			LOGGER.error("不是成都理想设备编号： {}", value);
			return value + "_error";
		}
	}

	private static String getPassword(String value) {

		if (Objects.isNullString(value)) {
			return "";
		}

		String prefix = "";

		int i = 10 - value.length();

		while (i-- > 0) {
			prefix += "0";
		}

		return prefix + value;
	}

	private static String getValue(Cell cell) {

		if (Objects.isNull(cell)) {
			return "";
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		if (Objects.isNullString(cell)) {
			return "";
		}
		return cell.getStringCellValue().trim();
	}

	private static String getPhoneValueString(Cell cell) {

		if (Objects.isNull(cell)) {
			return "";
		}
		if (Objects.isNullString(cell)) {
			return "";
		}

		try {
			BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
			return bd.toPlainString();
		} catch (Exception e) {
			// LOGGER.error(e.getMessage(), e);
			return getValue(cell);
		}
	}

	private static String getNumberValue(Cell cell) {

		if (Objects.isNull(cell)) {
			return "";
		}
		if (Objects.isNullString(cell)) {
			return "";
		}

		BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
		return bd.toPlainString();
	}

}
