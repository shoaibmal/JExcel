package maverick.jacob.utils;

import java.util.ArrayList;
import java.util.List;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JSheet {
private Dispatch sheet;
private Dispatch workbook;
	public JSheet(JWorkbook Workbook, String sheetName) {
		this.workbook = Workbook.getWorkbook();
		Dispatch.call(Workbook.getWorkbook(), "Activate");
		Dispatch sheets = Dispatch.call(Workbook.getWorkbook(),"Sheets").toDispatch();
		setSheet(Dispatch.call(sheets, "Item",new Variant(sheetName)).toDispatch());
	//	Dispatch.invoke(dispatchTarget, dispID, wFlags, oArg, uArgErr)(getSheet(),"Select");
		  Dispatch.call(sheet, "Select");
	}
	public JSheet(JWorkbook Workbook, int index) {
		this.workbook = Workbook.getWorkbook();
		Dispatch.call(Workbook.getWorkbook(), "Activate");
		setSheet( Dispatch.call(Workbook.getWorkbook(),"ActiveSheet").toDispatch());
	
	}
	public void setSheet(Dispatch sheet) {
		this.sheet = sheet;
	}
	public Dispatch getSheet() {
		return sheet;
	}
	public void select() {
		Dispatch.invoke(sheet, "Select", Dispatch.Get, new Object[] {}, new int[1]);
		   //Dispatch.invoke(sheet, "Select");	
	}
	public Dispatch getWorkbook() {
		return this.workbook;
	}
	public void setWorkbook(Dispatch workbook) {
		this.workbook = workbook;
	}
	public String getCellValue(String cellRef) {
		Dispatch.call(this.workbook,"Activate");
		Dispatch.call(this.sheet, "Select");
		  Dispatch resultcell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {cellRef}, new int[1]).toDispatch();
		   Variant cellValue = Dispatch.call(resultcell, "Value");
		return cellValue.toString();
	}
	public String getCellValue(int rowNum,int colNum){
		  Dispatch resultcell = Dispatch.invoke(sheet, "Cells", Dispatch.Get, new Object[] {rowNum,colNum}, new int[1]).toDispatch();
		  Variant cellValue = Dispatch.call(resultcell, "Value");
	return cellValue.toString();
	}
	public void sortColumn(String column1,String column2,String column3) {
		Dispatch.call(this.workbook,"Activate");
		Dispatch.call(this.sheet, "Select");
		Dispatch usedRange = Dispatch.call(this.sheet,"UsedRange").toDispatch();
		Dispatch usedRangeRow = Dispatch.call(usedRange,"Rows").toDispatch();
		Variant usedRangeRowCount = Dispatch.call(usedRangeRow,"Count");
		Dispatch usedRangeColumn = Dispatch.call(usedRange,"Columns").toDispatch();
		Variant usedRangeColumnCount = Dispatch.call(usedRangeColumn,"Count");
		System.out.println("Columns : "+usedRangeColumnCount);
		 Dispatch sortObj = Dispatch.call(sheet,"Sort").toDispatch();
	        Dispatch sortFieldObj = Dispatch.call(sortObj, "SortFields").toDispatch();
	        Dispatch.call(sortFieldObj, "Clear");
	        Object C1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {column1+"2:"+column1+usedRangeRowCount}, new int[1]).toDispatch();
	        Dispatch.call(sortFieldObj, "Add",new Variant(C1),new Variant(0),new Variant(1),new Variant(0)).toDispatch();
	        //xlSortOnValues-0
	        //xlAscending-1
	        //xlSortNormal-0
	        Object C2 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {column2+"2:"+column2+usedRangeRowCount}, new int[1]).toDispatch();
	        Dispatch.call(sortFieldObj, "Add",new Variant(C2),new Variant(0),new Variant(1),new Variant(0)).toDispatch();
	        Object C3 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {column3+"2:"+column3+usedRangeRowCount}, new int[1]).toDispatch();
	        Dispatch.call(sortFieldObj, "Add",new Variant(C3),new Variant(0),new Variant(1),new Variant(0)).toDispatch();
	       // Object UsedRange = Dispatch.call(sheet,"UsedRange"/*,new Variant("A1:BD29")*/).toDispatch();
	        Dispatch.call(sortObj, "SetRange", new Variant(usedRange));
	        Dispatch.put(sortObj,"Header",new Variant(1));//xlYes
	        Dispatch.put(sortObj,"MatchCase",new Variant(false));
	        Dispatch.put(sortObj,"Orientation",new Variant(1));//xlTopToBottom
	        Dispatch.put(sortObj,"SortMethod",new Variant(1));//xlPinYin
	        Dispatch.call(sortObj, "Apply");

		
	}
	public void sortColumn(String column1, String column2) {
		Dispatch.call(this.workbook,"Activate");
		Dispatch.call(this.sheet, "Select");
		Dispatch usedRange = Dispatch.call(this.sheet,"UsedRange").toDispatch();
		Dispatch usedRangeRow = Dispatch.call(usedRange,"Rows").toDispatch();
		Variant usedRangeRowCount = Dispatch.call(usedRangeRow,"Count");
		Dispatch usedRangeColumn = Dispatch.call(usedRange,"Columns").toDispatch();
		Variant usedRangeColumnCount = Dispatch.call(usedRangeColumn,"Count");
		System.out.println("Columns : "+usedRangeColumnCount);
		 Dispatch sortObj = Dispatch.call(sheet,"Sort").toDispatch();
	        Dispatch sortFieldObj = Dispatch.call(sortObj, "SortFields").toDispatch();
	        Dispatch.call(sortFieldObj, "Clear");
	        Object C1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {column1+"2:"+column1+usedRangeRowCount}, new int[1]).toDispatch();
	        Dispatch.call(sortFieldObj, "Add",new Variant(C1),new Variant(0),new Variant(1),new Variant(0)).toDispatch();
	        //xlSortOnValues-0
	        //xlAscending-1
	        //xlSortNormal-0
	        Object C2 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {column2+"2:"+column2+usedRangeRowCount}, new int[1]).toDispatch();
	        Dispatch.call(sortFieldObj, "Add",new Variant(C2),new Variant(0),new Variant(1),new Variant(0)).toDispatch();
	        Dispatch.call(sortObj, "SetRange", new Variant(usedRange));
	        Dispatch.put(sortObj,"Header",new Variant(1));//xlYes
	        Dispatch.put(sortObj,"MatchCase",new Variant(false));
	        Dispatch.put(sortObj,"Orientation",new Variant(1));//xlTopToBottom
	        Dispatch.put(sortObj,"SortMethod",new Variant(1));//xlPinYin
	        Dispatch.call(sortObj, "Apply");
	}
	public List<String> autoFilter(String srcHeaderName,String sortValue,String targetColumnName) {
		Dispatch.call(this.workbook,"Activate");
		Dispatch.call(this.sheet, "Select");
		Dispatch usedRange = Dispatch.call(this.sheet,"UsedRange").toDispatch();
		Dispatch usedRangeRow = Dispatch.call(usedRange,"Rows").toDispatch();
		Variant usedRangeRowCount = Dispatch.call(usedRangeRow,"Count");
		Dispatch usedRangeColumn = Dispatch.call(usedRange,"Columns").toDispatch();
		Variant usedRangeColumnCount = Dispatch.call(usedRangeColumn,"Count");
		int targetColumn = getColumnNumber(targetColumnName);
		int srcColumn = getColumnNumber(srcHeaderName);
		String usedRangeRef = "$A$1"+":"+getCellAddress(usedRangeRowCount.getInt(), usedRangeColumnCount.getInt());
		   Object C1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {"A1"}, new int[1]).toDispatch();
		  Dispatch.call((Dispatch) C1, "Select");
		  Dispatch.call((Dispatch) C1, "AutoFilter");
		  Dispatch C2= Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] {usedRangeRef}, new int[1]).toDispatch();
		  Dispatch.call(C2, "AutoFilter",new Variant(srcColumn),new Variant(sortValue));
		  List<String> targetValues = new ArrayList<String>();
		for(int i=2;i<= usedRangeRowCount.getInt();i++){
			if(!isRowHidden(i)){
				targetValues.add(getCellValue(i, targetColumn));
			}
		}
		return targetValues;
	}
	
	public int getColumnNumber(String headerName){
		Dispatch usedRange = Dispatch.call(this.sheet,"UsedRange").toDispatch();
		Dispatch usedRangeColumn = Dispatch.call(usedRange,"Columns").toDispatch();
		Variant usedRangeColumnCount = Dispatch.call(usedRangeColumn,"Count");
		for(int i=1;i<=usedRangeColumnCount.getInt();i++){
			if(getCellValue(1, i).trim().equals(headerName.trim())){
				return i;
			}
			}
		return -1;
		
	}
	public String getCellAddress(int rowNum,int colNum){
		Dispatch resultcell = Dispatch.invoke(sheet, "Cells", Dispatch.Get, new Object[] {rowNum,colNum}, new int[1]).toDispatch();
		Variant address = Dispatch.call(resultcell, "Address"); 
		return  address.toString();
	}
	public boolean isRowHidden(int rowNum){
		Dispatch row = Dispatch.invoke(sheet, "Rows", Dispatch.Get, new Object[] {rowNum}, new int[1]).toDispatch();
		Variant isRowHidden =  Dispatch.call( row, "Hidden");
		return isRowHidden.getBoolean();
	}
}
