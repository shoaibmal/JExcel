package maverick.jacob.utils;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.LibraryLoader;
import com.jacob.com.Variant;

public class JExcel {
	private ActiveXComponent excelInstance;
	public ActiveXComponent getExcelInstance() {
		return excelInstance;
	}

	public void setExcelInstance(ActiveXComponent excelInstance) {
		this.excelInstance = excelInstance;
	}

	public JExcel(){
	  init();
       this.excelInstance = new ActiveXComponent("Excel.Application");
       visbility(false);
       displayAlerts(true);
       
	}
	
	public JWorkbook openWorkbook(String fileName){
		JWorkbook workbook = new JWorkbook(this,fileName);
		return workbook;
	}
	public JWorkbook openWorkbook(File file){
		JWorkbook workbook = new JWorkbook(this,file);
		return workbook;
	}
	public void quit(){
		Dispatch.call(excelInstance, "Quit");
		 ComThread.Release();
	}
	public void displayAlerts(boolean condition){
		Dispatch.put(excelInstance, "DisplayAlerts", new Variant(condition));
	}
	public void visbility(boolean condition){
		Dispatch.put(excelInstance, "Visible", new Variant(condition));	
	}
	private void init(){
		 File file = new File("lib", "jacob-1.18-M2-x86.dll"); //path to the jacob dll
	       System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
      LibraryLoader.loadJacobLibrary();
	}
}
