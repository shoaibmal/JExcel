package maverick.jacob.utils;

import java.io.File;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JWorkbook {
	private Dispatch workbook;
	private Dispatch excel;
public JWorkbook(JExcel excel,String fileName){
	this.setExcel(excel.getExcelInstance());
    Dispatch workbooks = excel.getExcelInstance().getProperty("Workbooks").toDispatch();
   setWorkbook(Dispatch.call(workbooks, "Open", new Variant(fileName)).toDispatch());
}

public JWorkbook(JExcel excel,File file){
	this(excel,file.getAbsolutePath());
}
public void setWorkbook(Dispatch workbook) {
	this.workbook = workbook;
}
public Dispatch getWorkbook() {
	return workbook;
}
public JSheet getSheet(String sheetName){
	JSheet sheet = new JSheet(this,sheetName);
	return sheet;
}
public JSheet getSheetAt(int index){
	JSheet sheet = new JSheet(this,index);
	return sheet;
}
public void copySheet(JSheet srcSheet,JSheet targetSheet){
	 Dispatch.call(srcSheet.getWorkbook(),"Activate");
	 Dispatch range = Dispatch.call(srcSheet.getSheet(),"Range",new Variant("$A$1:$IV$65536")).toDispatch();
   //  Dispatch rows = Dispatch.call(range,"Rows").toDispatch();
   //  Variant count = Dispatch.call(rows, "Count");
     //System.out.println(count.getInt());
     Dispatch.call(range, "Copy");
    // Dispatch.call(targetSheet.getSheet(), "Select");
     Dispatch.call(targetSheet.getWorkbook(),"Activate");
     Dispatch a12 = Dispatch.invoke(targetSheet.getSheet(), "Range", Dispatch.Get, new Object[] {"A1"}, new int[1]).toDispatch();
    Dispatch.call(a12, "Select");
     Dispatch.call(targetSheet.getSheet(), "PasteSpecial",3);
     range = Dispatch.call(srcSheet.getSheet(),"Range",new Variant("A1")).toDispatch();
     Dispatch.call(range, "Copy");
}
public void copySheet(JSheet srcSheet,JSheet targetSheet,String startingCellRef){
	 Dispatch.call(srcSheet.getWorkbook(),"Activate");
	 Dispatch range = Dispatch.call(srcSheet.getSheet(),"Range",new Variant(startingCellRef+":$IV$65536")).toDispatch();
  //  Dispatch rows = Dispatch.call(range,"Rows").toDispatch();
  //  Variant count = Dispatch.call(rows, "Count");
    //System.out.println(count.getInt());
    Dispatch.call(range, "Copy");
   // Dispatch.call(targetSheet.getSheet(), "Select");
    Dispatch.call(targetSheet.getWorkbook(),"Activate");
    Dispatch a12 = Dispatch.invoke(targetSheet.getSheet(), "Range", Dispatch.Get, new Object[] {startingCellRef}, new int[1]).toDispatch();
   Dispatch.call(a12, "Select");
    Dispatch.call(targetSheet.getSheet(), "PasteSpecial",3);
    range = Dispatch.call(srcSheet.getSheet(),"Range",new Variant("A1")).toDispatch();
    Dispatch.call(range, "Copy");
}
public void saveAs(String fileName){
//	Dispatch.put(this.excel, "DisplayAlerts", new Variant(false));
	Dispatch.call(this.workbook,"SaveAs",new Variant(fileName));
    
}
public void save(){
//	Dispatch.put(this.excel, "DisplayAlerts", new Variant(false));
	Dispatch.call(this.workbook,"Save");
    
}
public void close(){
	Dispatch.call(this.workbook,"Close");
}

public void setExcel(Dispatch excel) {
	this.excel = excel;
}

public Dispatch getExcel() {
	return excel;
}
}
