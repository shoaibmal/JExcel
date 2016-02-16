package maverick.jacob.utils;

import java.io.File;

public class Test {
public static void main(String[] args) {
	
	File tc = new File("FULL FILE PATH");
	System.out.println("Exists : "+tc.exists());
	/*JExcel excel = new JExcel();
	excel.visbility(true);
	JWorkbook tcWorkbook =excel.openWorkbook("FULL FILE PATH");
	
//	eventActual.select();
	JWorkbook eventWorkbook = excel.openWorkbook("FULL FILE PATH");
	JSheet eventActual = tcWorkbook.getSheet("SHEET NAME");
	JSheet evnetSrcSheet = eventWorkbook.getSheetAt(1);
	tcWorkbook.copySheet(evnetSrcSheet, eventActual);
	excel.displayAlerts(false);
	JSheet resultSheet = tcWorkbook.getSheet("SHEETNAME");
	String status = resultSheet.getCellValue("A2");
	System.out.println("Resutl : "+status);
	tcWorkbook.saveAs("OUTPUT FULL FILE PATH");
	eventWorkbook.close();
	tcWorkbook.close();
	excel.quit();*/
}
}
