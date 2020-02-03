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
		
		/*Variabile che conterrà i nomi delle cartelle dei giochi rinominate*/
		String [] nomiCartelle = null; 
		
		/*Sara il nome del file del nuovo CM*/
		String nomeFoglioExc = ""; 
		
		/*Istanza della classe che gestisce tutto il ChangeManagement*/
		ChangeManagement CM;
		
		/*Vecchio File Excel */
		String vecchioCMExcel= "CM1_JAN_2020.xlsx";
		
		/*il mio foglio excel*/
		XSSFWorkbook workbook = null;
		
		/*Directory del nuovo CM, il mio obiettivo è quello di trovare tra i file che ci sono in questa cartella
		 * il file excel critical assets register così da poter copiare dei dati che mi servono*/
		File directoryNuovoCM = null; 
		
		DataOra dto = new DataOra();
		Scanner scanner = new Scanner(System.in);
				
		System.out.println("---------------------[Change Management]------------------------");
	    System.out.println(dto.DateStamp());
	    System.out.println("Path cartella Change Management ? ");
	    path = scanner.nextLine();
	   
	    if(path.length() > 0) {
	    	
	    	/* primo passo Change Management */
	    	System.out.println("---------- primo passo [Estrazione , Rinominazione , e Calcolo file sha1 giochi  ...] ------------");
	    	
	    	String oldCMPath = path +"\\CM1_JAN_2020";
	    	String newCMPath = path +"\\CM2_JAN_2020";
	    	
	    	RenameFolders rfNew = new RenameFolders(newCMPath); 
	    	System.out.println("Nuovo CM:");
	    	rfNew.renameFolders();
	    	
	    	RenameFolders rfOld = new RenameFolders(oldCMPath);
	    	System.out.println("\n\nVecchio CM:");
	    	rfOld.renameFolders();
	    	
	    	
	    	/* directory nuovo CM */
	    	directoryNuovoCM = new File(newCMPath); 
	    	
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
	    		
/*-----------------------------------------------[Caricamento dati Grezzi ]------------------------------------------------*/	    		
	    	    //caricamento dati in Grezzi
	    		System.out.print("Caricamento dati in Grezzi... ");
	    	    CM.grezzi(f,newCMPath);
	    	    System.out.print(" Fine\n");
/*---------------------------------------------------------[Fine]--------------------------------------------------------------*/	    	    
	    	    
	    	    
	    	    
/*-----------------------------------------------[Caricamento dati Checksums ]------------------------------------------------*/	    	    
	    		//caricamento dati Checksums 
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			File nuovoFile = new File(newCMPath+"\\"+nomeFoglioExc);
	    			File vecchioFile = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			System.out.print("Caricamento dati in Checksums ...");
	    			CM.checksums(nuovoFile,vecchioFile);
	    			//CM.checksums2(nuovoFile);
	    			System.out.print(" Fine\n");
	    		}else {
	    			 File nuovoFile = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx");
	    			 File vecchioFile = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			 System.out.print("Caricamento dati in Checksums ...");
	    			 CM.checksums(nuovoFile,vecchioFile);
	    			 //CM.checksums2(nuovoFile);
	    			 System.out.print(" Fine\n");
	    		}
/*---------------------------------------------------------[Fine]--------------------------------------------------------------*/	    	    
	    		
	    		
	    		
	    		
/*----------------------------------------[ Caricamento dati Appoggio Changed Games ]------------------------------------------*/
	    	    //caricamento dati Appoggio Changed Games
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			File vuoto = new File(newCMPath+"\\"+nomeFoglioExc); 
	    			File pieno = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			System.out.print("Caricamento dati in Appoggio Changed Games ...");
	    			 CM.appoggioChangedGames(vuoto,pieno);
	    			 System.out.print(" Fine\n");
	    		}else {
	    			File vuoto = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx"); 
	    			File pieno = new File(oldCMPath+"\\"+vecchioCMExcel);
	    			System.out.print("Caricamento dati in Appoggio Changed Games ...");
	    			 CM.appoggioChangedGames(vuoto,pieno); 
	    			 System.out.print(" Fine\n");
	    		}
/*---------------------------------------------------------[Fine]--------------------------------------------------------------*/
	    		
	    		
	    		
/*----------------------------------------------[caricamento dati Check EVO]-------------------------------------------------*/
	    		
	    		if(nomeFoglioExc.endsWith(".xlsx")) {
	    			
	    			System.out.print("Caricamento dati in Check EVO ...");
	    			File nuovoCM = new File(newCMPath+"\\"+nomeFoglioExc);
	    			File checkEvoFile = null; 
	    		    /* qui recupero tutti i nomi dei files presenti nella cartella del nuovo CM così troverò il file 
	    		     * Excel critical assets register */
	    			 String [] nomiFiles = directoryNuovoCM.list();
	    			 for(String s:nomiFiles) {
	    				 if(s.contains("critical assets register")) {
	    					 checkEvoFile = new File(newCMPath+"\\"+s);
	    				 }
	    			 }

	    			 CM.checkEVO(nuovoCM,checkEvoFile);
	    			 System.out.print(" Fine\n");
	    		}else {

	    			System.out.print("Caricamento dati in Check EVO ...");
	    			File nuovoCM = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx");
	    			File checkEvoFile = null; 
	    		    /* qui recupero tutti i nomi dei files presenti nella cartella del nuovo CM così troverò il file 
	    		     * Excel critical assets register */
	    			 String [] nomiFiles = directoryNuovoCM.list();
	    			 for(String s:nomiFiles) {
	    				 if(s.contains("critical assets register")) {
	    					 checkEvoFile = new File(newCMPath+"\\"+s);
	    				 }
	    			 }
	    			 
	    			 CM.checkEVO(nuovoCM,checkEvoFile);
	    			 System.out.print(" Fine\n");
	    		}
	    		
/*----------------------------------------------------------[Fine]-----------------------------------------------------------*/
	    	
	    		   		
/*------------------------------------------------[caricamento dati Report]----------------------------------------------------*/	    		
	    		
	    	   if( nomeFoglioExc.endsWith(".xlsx") ) {
	    		   
	    		   System.out.print("Caricamento dati in Report ...");
	    		   
	    			File nuovoCM = new File(newCMPath+"\\"+nomeFoglioExc);
	    			File reportFile = null; 
	    		    /* qui recupero tutti i nomi dei files presenti nella cartella del nuovo CM così troverò il file 
	    		     * Excel critical assets register */
	    			 String [] nomiFiles = directoryNuovoCM.list();
	    			 for(String s:nomiFiles) {
	    				 if(s.contains("Evolution Games Versions")) {
	    					 reportFile = new File(newCMPath+"\\"+s);
	    				 }
	    			 }
	    		     
	    			 CM.report(nuovoCM, reportFile);
	    			 
	    		   System.out.print(" Fine\n");
	    		   
	    	   }else {
	    		   
	    		   System.out.print("Caricamento dati in Report ...");
	    		   
	    			File nuovoCM = new File(newCMPath+"\\"+nomeFoglioExc+".xlsx");
	    			File reportFile = null; 
	    		    /* qui recupero tutti i nomi dei files presenti nella cartella del nuovo CM così troverò il file 
	    		     * Excel critical assets register */
	    			 String [] nomiFiles = directoryNuovoCM.list();
	    			 for(String s:nomiFiles) {
	    				 if( s.contains("Evolution Games Versions") ) {
	    					 reportFile = new File(newCMPath+"\\"+s);
	    				 }
	    			 }
	    		   
	    		   CM.report(nuovoCM, reportFile);
	    		   
	    		   System.out.print(" Fine\n");
	    		   
	    	   }
/*----------------------------------------------------------[Fine]-------------------------------------------------------------*/
	    		
	    	}else {  //nome foglio Excel non valido
	    		System.out.println("File non trovato! ");
	        }   	
	    	
	    }else {
	    	System.out.println("Path non corretto!");
	    }
	   
	}
	
}
