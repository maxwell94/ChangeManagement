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
		String vecchioCMExcel= "CM_DEC_2st_2019.xlsx";
		
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
	    	
	    	String oldCMPath = path +"\\CM2_DEC_2019";
	    	String newCMPath = path +"\\CM1_GEN_2020";
	    	
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
	    	rfNew.salvaSuFile(datiGrezzi, nomiCartelle, f);
	    	
	    	/* Fine primo passo */
	    	System.out.println("----[Fine primo passo ]-----\n");
	    	
	    	/* Secondo passo */
	    	System.out.println("---------- secondo passo [Lettura fileSha1.sha e caricamento dati Excel in Grezzi ...] ------------");
	    	
	    	System.out.print("Come vuoi chiamare il foglio excel per il nuovo CM ? :  ");
	    	nomeFoglioExc = scanner.nextLine();
	    	
	    	if(nomeFoglioExc.length() > 0) {
	    		
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			 CM = new ChangeManagement(newCMPath+"\\"+nomeFoglioExc); 
	    		}else {
	    			CM = new ChangeManagement(newCMPath+"\\"+nomeFoglioExc+".xlsx");
	    		}
	    		
	    	    //caricamento dati in Grezzi
	    	    CM.grezzi(f,newCMPath);
	    	    
	    	    //caricamento dati Appoggio Changed Games
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			File vuoto = new File(newCMPath+"\\"+nomeFoglioExc); 
	    			File pieno = new File(oldCMPath+"\\"+vecchioCMExcel); 
	    			 CM.appoggioChangedGames(vuoto,pieno);
	    			 System.out.println("\nDone ...");
	    		}else {
	    			File vuoto = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx"); 
	    			File pieno = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			 CM.appoggioChangedGames(vuoto,pieno); 
	    			 System.out.println("\nDone ..."); 
	    		}
	    		
	    		//caricamento dati Checksums 
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			File nuovoFile = new File(newCMPath+"\\"+nomeFoglioExc);
	    			File vecchioFile = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			CM.checksums1(nuovoFile,vecchioFile);
	    			//CM.checksums2(nuovoFile);
	    			System.out.println("\nDone ...");
	    		}else {
	    			 File nuovoFile = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx");
	    			 File vecchioFile = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			 CM.checksums1(nuovoFile,vecchioFile);
	    			 //CM.checksums2(nuovoFile);
	    			 System.out.println("\nDone ...");
	    		}
	    			    	    
	    	    
	    	}else {  //nome foglio Excel non valido
	    		System.out.println("File non trovato! ");
	        }   	
	    	
	    }else {
	    	System.out.println("Path non corretto!");
	    }
	   
	}
	
}
