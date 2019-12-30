import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.Timer;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws FileNotFoundException {
		
		String path = ""; 
		int rt = 0; 
		DataOra dto = new DataOra();
		Scanner scanner = new Scanner(System.in);
				
		System.out.println("---------------------[Change Management]------------------------");
	    System.out.println(dto.DateStamp());
	    System.out.println("Path cartella Change Management ? ");
	    path = scanner.nextLine();
	    
	    System.out.println("----[Estrazione e rinominazione giochi ...]-----");
	    
	    if(path.length() > 0) {
	    	
	    	String oldCMPath = path + "\\CM_NOV_1ST_2019";
	    	String newCMPath = path + "\\CM_NOV_2nd_2019";
	    	
	    	RenameFolders rfNew = new RenameFolders(newCMPath); 
	    	System.out.println("Nuovo CM:");
	    	rfNew.renameFolders();
	    	
	    	RenameFolders rfOld = new RenameFolders(oldCMPath);
	    	System.out.println("\n\nVecchio CM:");
	    	rfOld.renameFolders();
	    	
	    	
	    }else {
	    	System.out.println("Path non corretto!");
	    }
	    
	    System.out.println("----[Fine]-----");
	    
	}
	
}
