package excelfileconsolidator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.DefaultListModel;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JList;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excelfileconsolidator.fileprocessing.WorkbookCloner;
import excelfileconsolidator.filter.ExcelFilter;

public class ExcelFileConsolidatorLauncher {

	public static void main(String[] args) {

		// Obtain the path of the current directoy
		String sCarpAct = System.getProperty("user.dir");
		showGUI();
		System.out.println("END");
		System.exit(0);
	}

	public static void fileProcessor(String path) {
		System.out.println("READING EXCEL FILES AT: " + path);
		File currentFolder = new File(path);
		File[] filteredFiles = currentFolder.listFiles(new ExcelFilter());

		System.out.println(filteredFiles.length + " FILES WERE FOUND");
		// Creating master workbook
		Workbook wbMaster = new XSSFWorkbook();
		WorkbookCloner wbCloner = new WorkbookCloner(wbMaster);

		for (File currentFile : filteredFiles) {
			System.out.println("READING " + currentFile.getName());
			try {
				FileInputStream fileInputStream = new FileInputStream(currentFile);
				Workbook wbOrigin = new XSSFWorkbook(fileInputStream);
				wbCloner.clone(wbOrigin, currentFile.getName());

				wbOrigin.close();
				fileInputStream.close();
				
				System.out.println(currentFile.getName() + "WAS PROCESSED");
				
			} catch (FileNotFoundException fnfex) {
				fnfex.printStackTrace();
			} catch (IOException ioex) {
				ioex.printStackTrace();
			}
		}
		wbCloner.write(path);
	}

	public static void showGUI() {
		JFrame window = new JFrame("Excel File Consolidator");		
		
		window.setSize(640, 480);
		window.setVisible(true);
		
		JFileChooser fileExplorer = new JFileChooser();
		fileExplorer.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int result = fileExplorer.showOpenDialog(window);
		
		if(result == JFileChooser.APPROVE_OPTION) {
			File folder = fileExplorer.getSelectedFile();
			System.out.println(folder.getAbsolutePath());
			fileProcessor(folder.getAbsolutePath());
		}
	}
}