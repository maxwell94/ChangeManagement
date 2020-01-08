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
		String nomeFoglioExc = ""; 
		ChangeManagement CM;
		
		//il mio foglio excel
		XSSFWorkbook workbook;
		
		DataOra dto = new DataOra();
		Scanner scanner = new Scanner(System.in);
				
		System.out.println("---------------------[Change Management]------------------------");
	    System.out.println(dto.DateStamp());
	    System.out.println("Path cartella Change Management ? ");
	    path = scanner.nextLine();
	   
	    if(path.length() > 0) {
	    	
	    	/* primo passo Change Management */
	    	System.out.println("---------- primo passo [Estrazione , Rinominazione , e Calcolo file sha1 giochi  ...] ------------");
	    	
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
	    	nomiCartelle = rfNew.saveFolderNames(newCMPath);
	    	rfNew.salvaSuFile(datiGrezzi, nomiCartelle,f);
	    	
	    	/* Fine primo passo */
	    	System.out.println("----[Fine primo passo ]-----\n");
	    	
	    	System.out.println("---------- secondo passo [Lettura fileSha1.sha e caricamento dati Excel in Grezzi ...] ------------");
	    	
	    	System.out.println("Come vorresti chiamare il tuo foglio excel : ");
	    	nomeFoglioExc = scanner.nextLine();
	    	
	    	if(nomeFoglioExc.length() > 0) {
	    		
	    	    CM = new ChangeManagement(nomeFoglioExc);
	    	    
	    	    //caricamento dati in grezzi
	    	    CM.grezzi(f);
	    	    
	    	}else {  //nome foglio Excel non valido
	    		
	    	}   	
	    	
	   
	    }else {
	    	System.out.println("Path non corretto!");
	    }
	   
	    
	}
	
}
