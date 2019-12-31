import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.Timer;
import java.util.TreeMap;
import java.util.concurrent.locks.StampedLock;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws NoSuchAlgorithmException, IOException {
		
		String path = ""; 
		int rt = 0; 
		ArrayList<String> datiGrezzi = new ArrayList<String>(); 
		String [] nomiCartelle = null; 
		
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
	    	
	    	//recupero tutti i dati da scrivere nel file che contiene le info di grezzi
	    	datiGrezzi = rfNew.getAllInfo();
	    	
	    	//creo un file
	    	File f = new File(newCMPath+"\\filesSha1.sha");
	    	//e scrivo questi dati su file
	    	rfNew.salvaSuFile(datiGrezzi, f);
	    	nomiCartelle = rfNew.saveFolderNames(newCMPath);
	    	rfNew.stampaArray(nomiCartelle);
	    	
	    	
	    }else {
	    	System.out.println("Path non corretto!");
	    }
	    
	    System.out.println("----[Fine]-----");
	    
	}
	
}
