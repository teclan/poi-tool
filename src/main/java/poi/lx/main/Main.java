package poi.lx.main;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import poi.lx.db.DatabaseServer;
import poi.lx.utils.ExcelUtils;

public class Main {
	private static final Logger LOGGER = LoggerFactory.getLogger(Main.class);

	public static void main(String[] args) {

		String filePath1 = "理想用户20180112\\2016新装客户资料汇总.xlsx";
		String filePath2 = "理想用户20180112\\2017新装客户资料汇总.xlsx";
		String filePath3 = "理想用户20180112\\2018新装客户资料汇总.xlsx";
		String filePath4 = "理想用户20180112\\2018.1-10未归档.xlsx";

		DatabaseServer.openDatabase();

		DatabaseServer.getDb().exec("delete from imm_devinfo");
		DatabaseServer.getDb().exec("delete from imm_netnvrattr");
		DatabaseServer.getDb().exec("delete from imm_camera");
		LOGGER.info("\n====================== 开始解析文件:{}\n", filePath1);
		ExcelUtils.analyze(filePath1);
		LOGGER.info("\n====================== 开始解析文件:{}\n", filePath2);
		ExcelUtils.analyze(filePath2);
		LOGGER.info("\n====================== 开始解析文件:{}\n", filePath3);
		ExcelUtils.analyze(filePath3);
		LOGGER.info("\n====================== 开始解析文件:{}\n", filePath4);
		ExcelUtils.analyzeNOArch(filePath4);
		DatabaseServer.closeDatabase();

	}
}
