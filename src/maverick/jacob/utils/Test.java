package maverick.jacob.utils;

import java.io.File;

public class Test {
public static void main(String[] args) {
	
	File tc = new File("C:\\ICAM\\QCDownloadedTestCases\\CM_CC_ASVN=0_001.xls");
	System.out.println("Exists : "+tc.exists());
	/*JExcel excel = new JExcel();
	excel.visbility(true);
	JWorkbook tcWorkbook =excel.openWorkbook("C:\\Users\\mm821700\\Desktop\\Misc\\CM_CC_ASVN=0_002.xls");
	
//	eventActual.select();
	JWorkbook eventWorkbook = excel.openWorkbook("C:\\Users\\mm821700\\Desktop\\Misc\\143073_1_2.0_2500_10.0_70.0_500.0_0_Event.csv");
	JSheet eventActual = tcWorkbook.getSheet("EventHistory-Actual");
	JSheet evnetSrcSheet = eventWorkbook.getSheetAt(1);
	tcWorkbook.copySheet(evnetSrcSheet, eventActual);
	excel.displayAlerts(false);
	JSheet resultSheet = tcWorkbook.getSheet("Results");
	String status = resultSheet.getCellValue("A2");
	System.out.println("Resutl : "+status);
	tcWorkbook.saveAs("C:\\Users\\mm821700\\Desktop\\Misc\\CM_CC_ASVN=0_002_TC.xls");
	eventWorkbook.close();
	tcWorkbook.close();
	excel.quit();*/
}
}
