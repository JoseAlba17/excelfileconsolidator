package excelfileconsolidator.filter;

import java.io.File;
import java.io.FileFilter;

public class ExcelFilter implements FileFilter {

	public boolean accept(File pathname) {
		
		return pathname.getName().endsWith(".xls") || 
				pathname.getName().endsWith(".xlsm") ||
				pathname.getName().endsWith(".xlsb") ||
				pathname.getName().endsWith(".xlsx");
	
	}
}
