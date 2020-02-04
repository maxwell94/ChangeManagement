import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.examples.CreateCell;

/**
 * Classe che gestisce tutto il Change Management. 
 * Ciascun metodo ha un ruolo ben preciso , per esempio:
 * il metodo "appoggioChangedGames()" andra nel tab "Appoggio Changed Games" 
 * e gestira tutto il lavoro che dovrà fare in "Appoggio Changed Games" e restituirà
 * il file aggiornato,
 * il metodo "checkSums()" andrà nel tab "Checksums" e gestira tutto il lavoro che dovrà
 * fare in checksum e restituira il file aggiornato e così via.
 * Author : Dedou Maximin */

public class ChangeManagement {

	/* Il foglio excel che contiene tutto il Change Management */
	private XSSFWorkbook workbook;
	
	/* I vari tabs del foglio excel */
	private XSSFSheet [] tabs;
	
	private String nomeFoglioExcel;
	
	/* controllare se la creazione del foglio Ã¨ andata buon fine */
	private int rt ; 
	
	ArrayList<String> giochiPathGrezzi ;
	
	/*---------- [da cancellare] -----------*/
	ArrayList<String> checksumColumn1 ;
	ArrayList<String> checksumColumn2 ;
	ArrayList<String> checksumColumn3 ;
	ArrayList<String> checksumColumn4 ;
	ArrayList<String> checksumColumn5Sha1; 
	/*--------------------------------------*/
	
	ArrayList<String> appoChgGamesColumn3; 
	ArrayList<String> appoChgGamesColumn4;
	ArrayList<String> appoChgGamesColumn5;
	ArrayList<String> checksumsColumn1;
	ArrayList<String> checksumsColumn2;
	ArrayList<String> checksumsColumn5;
	
	ArrayList<String> EGVGameName;
	ArrayList<String> EGVGameType;
	ArrayList<String> EGVPlatformVersion;
	ArrayList<String> EGVGameVersion;
	ArrayList<String> gameVersionsGameName; 
	
	
	String Sha1PrepareChecksumsGrezzi;
	String RngService; 
	
	static int nDatiGrezzi;
	static int nDatiChecksums; 

	/*Costruttore*/
	public ChangeManagement(String nomeFoglioExcel) {
		
		this.workbook = new XSSFWorkbook();
		this.tabs = new XSSFSheet [8];
		this.nomeFoglioExcel = nomeFoglioExcel ; 
		
		/*creo i vari tabs del file excel per il Change Management*/
		this.tabs[0] = this.workbook.createSheet("Grezzi"); // crea un tab Grezzi
		this.tabs[1] = this.workbook.createSheet("Checksums"); // crea un tab Checksums
		this.tabs[2] = this.workbook.createSheet("Report"); // crea un tab Report
		this.tabs[3] = this.workbook.createSheet("Game Versions"); // crea un tab Game Versions
		this.tabs[4] = this.workbook.createSheet("Changed Games"); // crea un tab Changed Games
		this.tabs[5] = this.workbook.createSheet("Appoggio Changed Games"); // crea un tab Appoggio Changed Games
		this.tabs[6] = this.workbook.createSheet("Check EVO"); // crea un tab Check Evo
		this.tabs[7] = this.workbook.createSheet("Description"); // crea un tab Description
		
		this.giochiPathGrezzi = new ArrayList<String>() ; 
		
		this.checksumColumn1 = new ArrayList<String>() ;
		this.checksumColumn2 = new ArrayList<String>() ;
		this.checksumColumn3 = new ArrayList<String>() ;
		this.checksumColumn4 = new ArrayList<String>() ;
		this.checksumColumn5Sha1 = new ArrayList<String>() ; 
		this.Sha1PrepareChecksumsGrezzi = "" ; 
		this.RngService = ""; 
		
		
		this.appoChgGamesColumn3 = new ArrayList<String>() ; 
		this.appoChgGamesColumn4 = new ArrayList<String>() ; 
		this.appoChgGamesColumn5 = new ArrayList<String>() ; 
		this.checksumsColumn1 = new ArrayList<String>() ; 
		this.checksumsColumn2 = new ArrayList<String>() ; 
		this.checksumsColumn5 = new ArrayList<String>() ; 
		
		this.EGVGameName = new ArrayList<String>() ; 
		this.EGVGameType = new ArrayList<String>() ; 
		this.EGVPlatformVersion = new ArrayList<String>() ; 
		this.EGVGameVersion = new ArrayList<String>() ; 
		this.gameVersionsGameName = new ArrayList<String>() ; 
		

		nDatiGrezzi = 0 ; 
		nDatiChecksums = 0; 
		
		//creo fisicamente il foglio excel
		rt = generaFoglioExcel(nomeFoglioExcel);
		
	}
	
	
	/* --------------------------------------------[generaFoglioExcel]---------------------------------------------------*/
		/* Metodo che genera un foglio excel che il Change Management nella cartella di lavoro 
	 * (in futuro si potra modificare cosi da scegliere il percorso in cui si vorrebbe salvarlo) */
	public int generaFoglioExcel(String nome) {

		int retcode = 0;
		
		try {
			 
			 if(nome != null) {
				 
				 File f = new File(nome);
				 if(!f.exists()) { // se non esiste lo crea 
					 
					 FileOutputStream out = new FileOutputStream(new File(nome));
					 this.workbook.write(out);
					 out.close();
					 retcode = 0;
				 }else {
					 
					 retcode = -1;
				 }

			 }else {
				 System.out.println("Specificare il nome del file!");
				 retcode = -1;
			 }
			 
		}catch(Exception ex) { ex.printStackTrace(); }
		
		return retcode;
		
	}
	
	/*--------------------------------------------------[Fine]------------------------------------------------------------*/
	
	
	
	/* ---------------------------------------------[leggiCheckColumn5Sha1]---------------------------------------------*/
	
	public void leggiChecksumsColumn5Sha1(File f) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( f ); //file nuovo CM
			XSSFWorkbook w = new XSSFWorkbook(mioFile);
			XSSFSheet wSheet = w.getSheetAt(1);
			
			FormulaEvaluator fe = w.getCreationHelper().createFormulaEvaluator();
			
			Iterator<Row> rowIterator = wSheet.iterator() ;
			
			//System.out.println("DEBUG1");
			
    		for(int i=1; i<nDatiChecksums; i++) {
    			
        		Cell c5 = wSheet.getRow(i).getCell(10); //column5
        		
        		
        		CellValue cellValueColumn5 = fe.evaluate(c5);
        		
        		/*Se uguale a null controllo se è la riga corrispondente a quella del file prepare_checksums.py */
        		if(cellValueColumn5.getStringValue() == null) {
        			
        			if(checksumColumn4.get(i).contains("prepare_checksums.py")) {
        			   /* Aggiungo lo sha1 di "prepare_checksums.py" al mio vettore così da recuperarlo 
        			    * in Appoggio Changed Games per la scrittura */	
        			   checksumColumn5Sha1.add(Sha1PrepareChecksumsGrezzi);	
        			}
    				//System.out.println("i = "+i+" str vuota");
    			}
        		
        	  switch(cellValueColumn5.getCellType()) {
        		   
        		case STRING :
        			checksumColumn5Sha1.add(cellValueColumn5.getStringValue());	
        		break;
        		
        	  }
    			
    		}
			
		}catch(IOException ex) { ex.printStackTrace(); }
		
	}
	
/*--------------------------------------------------[Fine]-----------------------------------------------------------*/
	
	
/* ---------------------------------------------[leggiAppoggioChgGamesVCM]---------------------------------------------*/
	
	public void leggiAppoggioChgGamesVCM(XSSFSheet mySheet) {
		
		int rowNum = 0; 
		Iterator<Row> rowIterator = mySheet.iterator();
		
		/* Salto la prima riga perché è il titolo */
		if(rowIterator.hasNext()) {
			rowIterator.next() ; 
		}
		
		while(rowIterator.hasNext()) {
			
			Row rigaLetta = rowIterator.next() ; 
			
			Iterator<Cell> cellIterator = rigaLetta.cellIterator() ;
			
			int cellNum = 0; 
			
			while(cellIterator.hasNext()) {
				
				Cell cellaLetta = cellIterator.next() ; 
				
				if(cellNum == 0) {
					appoChgGamesColumn3.add( cellaLetta.getStringCellValue() );
					//System.out.print(cellaLetta.getStringCellValue()+"  ");
					
				}else if(cellNum == 1) {
					appoChgGamesColumn4.add( cellaLetta.getStringCellValue() );
					//System.out.print(cellaLetta.getStringCellValue()+"  ");
					
				}else if(cellNum == 2) {
					appoChgGamesColumn5.add( cellaLetta.getStringCellValue() );
					//System.out.print(cellaLetta.getStringCellValue()+"\n");
				}
				
				
				cellNum ++;
			}
			
			rowNum ++;
		}
	}
	
/*--------------------------------------------------------[Fine]---------------------------------------------------------*/
	
	
	
	
/* ------------------------------------------------------[leggiChecksumsNCM]--------------------------------------------------*/
	
	public void leggiChecksumsNCM(XSSFSheet mySheet, FormulaEvaluator fe) {
		
		int rowNum = 0;
		String str = "";
		
		for(Row row: mySheet) {
			
			for(int cn=0; cn < row.getLastCellNum(); cn ++) {
				
				Cell cell = row.getCell(cn ,MissingCellPolicy.CREATE_NULL_AS_BLANK) ;
				
				if(cn == 5 && rowNum > 0 && !cell.toString().isEmpty()) {   //column1 
					//System.out.print(cell.toString()+"   ") ;
					checksumsColumn1.add(cell.toString()); 
					
				}else if( cn == 6  && rowNum > 0 && !cell.toString().isEmpty()) {   //column2
					checksumsColumn2.add(cell.toString()); 
					//System.out.print(cell.toString()+"   ") ;
					str = cell.toString() ; 
					
				}else if( cn == 10  && rowNum > 0 && !cell.toString().isEmpty() ) {  //column5
					
					CellValue cellValue = fe.evaluate(cell);
					
					if( cellValue.getStringValue() == null) {
						
						if( str.equals("prepare_checksums.py") ) {
							checksumsColumn5.add(Sha1PrepareChecksumsGrezzi);
						}
						//System.out.println("Sha1 = null");
					}else {
						
						switch( cellValue.getCellType() ) {
					     
						   case STRING:
							   checksumsColumn5.add(cellValue.getStringValue());
						       //System.out.print(cellValue.getStringValue()+"\n") ;
						   break; 
						}
					}
					
				}
			}
			
			rowNum ++;
		}
	}

/*-------------------------------------------------------------[Fine]---------------------------------------------------------*/
	
	
/* ---------------------------------------------[appoggioChangedGames]---------------------------------------------*/

	/* Metodo che lavora all'interno del tab Appoggio Changed Games 
	 * getRichStringCellValue().toString();*/
	public void appoggioChangedGames(File f , File f2)  {
		
		try {
			
			/*File nuovo CM*/
			FileInputStream fileNuovoCM = new FileInputStream( f ); 
			
			/*File vecchio CM*/
			FileInputStream fileVecchioCM = new FileInputStream(f2) ;  
			
			/* Foglio nuovo CM */
			XSSFWorkbook workbook_ncm = new XSSFWorkbook(fileNuovoCM);
			
			/*Foglio vecchio CM*/
			XSSFWorkbook workbook_vcm = new XSSFWorkbook(fileVecchioCM) ; 
			
		    /*vado nel tab desiderato Appoggio Changed Games in entrambi i fogli */
			XSSFSheet tabAppoggioNCM = workbook_ncm.getSheetAt(5);
			XSSFSheet tabAppoggioVCM = workbook_vcm.getSheetAt(5);
			
			/*Vado nel tab Checksums nel file del nuovo CM */
			XSSFSheet tabChecksumsNCM = workbook_ncm.getSheetAt(1);
			
			int numRowTabNCM = 0; 
			
			/*pulisco il tab Appoggio Changed Games del nuovo CM prima dell'inserimento di nuovi dati*/
			deleteSheetAllContent(tabAppoggioNCM);
			
			/*Dopo aver pulito il tab Appoggio Changed Games del nuovo CM inserisco il titolo */
			Row rigaTitolo = tabAppoggioNCM.createRow(numRowTabNCM);
			
			rigaTitolo.createCell(0).setCellValue("C_Game");
			rigaTitolo.createCell(1).setCellValue("C_File");
			rigaTitolo.createCell(2).setCellValue("C_Sha1");
			rigaTitolo.createCell(3).setCellValue("O_Game");
			rigaTitolo.createCell(4).setCellValue("O_File");
			rigaTitolo.createCell(5).setCellValue("O_Sha1");
			rigaTitolo.createCell(6).setCellValue("Game");
			rigaTitolo.createCell(7).setCellValue("File");
			rigaTitolo.createCell(8).setCellValue("Sha");
			numRowTabNCM ++;
			
			/*Leggo il tab Appoggio Changed Games del vecchio CM e memorizzo i dati delle prime tre colonne dentro 
			 * tre ArrayList diversi */
			leggiAppoggioChgGamesVCM(tabAppoggioVCM);
			
			/*Leggo il tab Checksums del nuovo CM e memorizzo i dati delle colonne column1 , column2 , column5
			 *dentro tre ArrayList diversi */
			FormulaEvaluator fe = workbook_ncm.getCreationHelper().createFormulaEvaluator() ;  
			leggiChecksumsNCM(tabChecksumsNCM , fe);
			
			int iGame = 2; 
			int iFile = 2; 
			int iSha1 = 2; 
			
			
			
//			System.out.println("size appoChgGamesColumn3 = "+appoChgGamesColumn3.size());
//			System.out.println("size appoChgGamesColumn4 = "+appoChgGamesColumn4.size());
//			System.out.println("size appoChgGamesColumn5 = "+appoChgGamesColumn5.size());
//			
//			System.out.println("size checksumsColumn1 = "+checksumsColumn1.size());
//			System.out.println("size checksumsColumn2 = "+checksumsColumn2.size());
//			System.out.println("size checksumsColumn5 = "+checksumsColumn5.size());
			
			
			/*Scrivo tutti i dati letti e memorizzati nel tab Appoggio Changed Games del nuovo CM */
			for(int i=0; i<appoChgGamesColumn3.size(); i++) {
				
				/* Creo una riga nel tab Appoggio Changed Games del nuovo CM per ogni info letta */
				Row rigaTabNCM = tabAppoggioNCM.createRow(numRowTabNCM);
				
				/* E ora creo tre celle e inserisco le informazioni memorizzate */
				rigaTabNCM.createCell(0).setCellValue( checksumsColumn1.get(i) );
				rigaTabNCM.createCell(1).setCellValue( checksumsColumn2.get(i) );
				rigaTabNCM.createCell(2).setCellValue( checksumsColumn5.get(i) );
				rigaTabNCM.createCell(3).setCellValue( appoChgGamesColumn3.get(i) );
				rigaTabNCM.createCell(4).setCellValue( appoChgGamesColumn4.get(i) );
				rigaTabNCM.createCell(5).setCellValue( appoChgGamesColumn5.get(i) );
				
				String formulaIGAME = "IF(A"+iGame+"=D"+iGame+",TRUE)"; 
				String formulaIFILE = "IF(B"+iFile+"=E"+iFile+",TRUE)";
				String formulaISHA1 = "IF(C"+iSha1+"=F"+iSha1+",TRUE)";
				
				rigaTabNCM.createCell(6).setCellFormula(formulaIGAME);
				rigaTabNCM.createCell(7).setCellFormula(formulaIFILE);
				rigaTabNCM.createCell(8).setCellFormula(formulaISHA1);
				
				iGame ++;
				iFile ++;
				iSha1 ++;
				
				numRowTabNCM ++;
			}
			
//			for(int i=0;i<appoChgGamesColumn3.size(); i++) {
//				System.out.println( checksumsColumn2.get(i) +" <-----> "+appoChgGamesColumn4.get(i)) ;
//			}
//			System.out.println();
			
			/*Ora leggo e memorizzo i dati della colonna 5 del tab Checksums del file del nuovo CM */
			//leggiChecksumsColumn5Sha1(f);
			
			/*Scrivo ora i dati letti e memorizzati nel tab Appoggio Changed games del nuovo CM */
			
			/* setto a 1 il contatore delle righe che vengono create perché scrivere nelle altre tre colonne 
			 * partendo dalla prima riga */
			numRowTabNCM = 1; 
			
			
			
			fileNuovoCM.close(); 
			fileVecchioCM.close();
			
			//aggiorna il file che era vuoto
			FileOutputStream out = new FileOutputStream(f);
			workbook_ncm.write(out);
			out.close();

		}catch(Exception ex) { ex.printStackTrace(); }
	}

/*--------------------------------------------------[Fine]----------------------------------------------------------*/
	
	
	
	
	
/* ---------------------------------------------[sostituisci]--------------------------------------------------------*/
	/* metodo che prende una stringa sostituisce '\' con '/' e ritorna un array */
	public static char [] sostituisci (String str) {
		
		char [] myStrChar = null; 
		
		if( str.length() > 0 ) {
			
			myStrChar = str.toCharArray();
			
			for(int i=0; i<myStrChar.length; i++) {
				
				if(  myStrChar[i] == '\\' ) {
					myStrChar[i] = '/'; 
				}
			}
		}
		
		return myStrChar;
	}
	
/*-------------------------------------------------[Fine]------------------------------------------------------------*/
	
	
	
	
/* ---------------------------------------------[deleteSheetAllContent]---------------------------------------------*/
	
	/* Metodo che cancella tutto il contenuto di un preciso tab dentro un file Excel */
	public void deleteSheetAllContent(XSSFSheet desiredSheet) {
		
		Iterator<Row> rowIterator = desiredSheet.iterator();
		
        while (rowIterator.hasNext()) 
        {
            Row row = rowIterator.next();
            //For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();
             
            while (cellIterator.hasNext()) 
            {
                Cell cell = cellIterator.next();
              
                switch (cell.getCellType()) 
                {
                    case NUMERIC:
                    	cell.setCellType(CellType.BLANK);
                        break;
                    case STRING:
                    	cell.setCellType(CellType.BLANK);
                        break;
                    case FORMULA:
                    	cell.setCellType(CellType.BLANK);
                        break;    
                }
            }
        }
	}

/*-----------------------------------------------[Fine]--------------------------------------------------------------*/
	
	
	
	
/* ---------------------------------------------[grezzi]------------------------------------------------------------*/
	
	/* Metodo che lavora all'interno del tab Grezzi , praticamente 
	 * legge il file degli sha e lo carica in excel in grezzi*/
	public void grezzi(File f,String pth) throws IOException, NoSuchAlgorithmException {
		
		if(f.exists()) {
			
			BufferedReader br = new BufferedReader(new FileReader(f)); 
	        String st; 
	        String sha1CheckSum = ""; 
	        String nomeCheckSum = ""; 
	        char [] appoggio = null; 
	        
	        /* apro il mio foglio excel */
			//FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
			//XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
			
			//vado nel tab desiderato che in questo caso è Grezzi 
			XSSFSheet desiredSheet = workbook.getSheetAt(0);
			//System.out.print("Aggiornamento tab: "+desiredSheet.getSheetName()+"...");
			
			/* cancello tutti i dati che ci sono già nel tab grezzi */
			deleteSheetAllContent(desiredSheet);
            
			
	        /* File prepare_checksums.py da aggiungere alla lista grezzi*/
	        String pathFileChecksum = pth + "\\prepare_checksums.py";
	        File fileChecksum = new File(pth);
	        if(fileChecksum.exists()) {
	        	FileSha1 fs1 = new FileSha1() ;
	        	sha1CheckSum = fs1.sha1Code(pathFileChecksum) ; 
	        	nomeCheckSum = "/prepare_checksums.py"; 
	        	//System.out.println(fs1.sha1Code(pathFileChecksum)+" "+fileChecksum.getName());
	        }else {
	        	System.out.println("File prepare_checksums.py non esistente!");
	        }
			
			
			
			/* e ora copio i nuovi dati presi dal file sha1 */
			int rowNum = 0; 
			
	        while( (st = br.readLine() ) != null ) {
	        	
	        	appoggio = sostituisci(st);
	        	String dati = new String(appoggio);
	        	String sha1 = dati.substring(0,40);
	        	String path = dati.substring(41);
	        	
	        	Row row = desiredSheet.createRow(rowNum);
	        	nDatiGrezzi++;
	        	
	        	if(path.endsWith("RngService.class")) {
	        		RngService = sha1 ;
	        	}
	        	
	        	int cellNum = 0; 
	        	for(int i=0;i<2; i++) {
	        		
	        		Cell cell = row.createCell(i);
	        		if(i == 0) {
	        			cell.setCellValue(sha1);
	        		}else {
	        			
	        			/* tolgo dal path il nome della cartella principale del gioco */
	        			//System.out.println("path: "+path);
	        			//int posPrimoSlash = path.indexOf("/");
	        			//String nuovoPath = path.substring(posPrimoSlash+1); 	        			
	        			cell.setCellValue(path);
	        			giochiPathGrezzi.add(path); 
	        		}
	        	}
	        	 
	        	rowNum ++; 
	        }
	        
	        /* add file prepare_checksums.py*/
	        Row r = desiredSheet.createRow(rowNum++);
	        for(int i=0;i<2;i++) {
	        	Cell c = r.createCell(i);
        		if(i == 0) {
        			Sha1PrepareChecksumsGrezzi = sha1CheckSum.toLowerCase(); 
        			c.setCellValue(sha1CheckSum.toLowerCase());
        		}else {
        			c.setCellValue(nomeCheckSum);
        			
        			giochiPathGrezzi.add(nomeCheckSum);
        		}
	        }
	        
	        rowNum++;
	        nDatiGrezzi++; 
	        
			//aggiorna il foglio
			FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			//System.out.print(" Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
		}else {
			System.out.println("File non trovato! ");
		}
		
	}
	
/*-----------------------------------------------[Fine]--------------------------------------------------------------*/
	
	
	
	
	
/* ---------------------------------------------[cercaPaths]---------------------------------------------------------*/
	public int cercaPaths(String s) {
		
		int trovato = 0; 
		for(String str: giochiPathGrezzi) {
			if(s.equals(str)) {
				trovato = 1 ; 
			}
		}
		
		return trovato ; 
	}
	
/*-------------------------------------------------[Fine]------------------------------------------------------------*/
	
	
	
/* ---------------------------------------------[caricaDescription]-----------------------------------------------------*/	
	/* Metodo che lavora dentro il tab Description */
	public void caricaDescription(XSSFSheet descriptionSheet, File v) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( v ) ;
			XSSFWorkbook w = new XSSFWorkbook(mioFile); 
			
			XSSFSheet desiredSheet = w.getSheetAt(7); 
			
			Iterator<Row> rowIterator = desiredSheet.iterator(); 
			
			int rowNum = 0 ; 
		
			
			/* Nel sheet da riempire creo prima il titolo */
			Row header = descriptionSheet.createRow(rowNum);
			
		    header.createCell(0).setCellValue("Filename");
		    header.createCell(1).setCellValue("Path");
		    header.createCell(2).setCellValue("Description");
			
			/* Invece nel sheet da cui devo copiare le cose è la prima riga la salto perché è il titolo */
			if(rowIterator.hasNext()) {
				rowIterator.next();
				rowNum ++;
			}
			 
			
			XSSFSheet sheetGrezzi = w.getSheetAt(0); 
			
			while(rowIterator.hasNext()) {
				
				Row row = rowIterator.next() ; 
				
				Row row2 = descriptionSheet.createRow(rowNum) ; 
				
				Iterator<Cell> cellIterator = row.cellIterator() ; 
				
				int cellNum = 0;
							
				
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					
					Cell c = row2.createCell(cellNum) ; 
					
					switch(cell.getCellType()) {
					
						case STRING:
                            
							/* Cercherò in Grezzi il path corrente, se non lo trovo toglierò il nome del folder principale
							 * dal path ed eseguirò di nuovo la ricerca. è un modo per sapere devo modificare il path o no*/
							
							c.setCellValue(cell.getStringCellValue());
							
							break; 
					}
					
					cellNum ++; 
				}
				
				rowNum ++; 
			}
			
			
		}catch(IOException ex) {ex.printStackTrace();}
		 
	}
	
/*----------------------------------------------------[Fine]------------------------------------------------------------*/
	
	
	
	
/* --------------------------------------------------[leggiVCM]-------------------------------------------------------*/	
	/* Metodo che va a leggere il foglio excel del vecchio CM e memorizza la column1 , column2 , column3 , 
	 * column4 dentro degli arraylist */
	
	public void leggiVCM(File f) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( f );	
			XSSFWorkbook w = new XSSFWorkbook(mioFile);
			
			//FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator() ; 
			
			XSSFSheet wSheet = w.getSheetAt(1); //vado nel tab Checksums
	
			
			int rowNum = 0; 
			
			for(Row row: wSheet) {
				
				
				for(int cn = 0; cn < row.getLastCellNum(); cn++) {
					
					Cell cell = row.getCell(cn ,MissingCellPolicy.CREATE_NULL_AS_BLANK) ;
					if(cn == 5) {
						//System.out.print(cell.toString()+"   ") ;
						checksumColumn1.add(cell.toString()); 
						
					}else if( cn == 6) {
						checksumColumn2.add(cell.toString());
						//System.out.print(cell.toString()+"   ") ;
						
					}else if(cn == 7) {
						checksumColumn3.add(cell.toString());
						//System.out.print(cell.toString()+"   ") ;
						
					}else if(cn == 8) {
						
						/* Devo togliere il nome del folder principale però ricordarsi che dal CM1_JAN_2020 
						 * in poi togliere questa parte */
						checksumColumn4.add(cell.toString());
						
						//System.out.print(cell.toString()) ;
					}
					
				}
				
			}
			
			
		}catch(IOException ex) { ex.printStackTrace(); }
	
	}
	
/*----------------------------------------------------[Fine]-----------------------------------------------------------*/
	
	


/* -----------------------------------------------[deleteEmptyRows]---------------------------------------------------------*/
	/* Metodo che le celle vuote dentro un tab */
	public static void deleteEmptyRows(XSSFSheet checksumsSheet) {
		
		int indexRow = 0; 
		
		//System.out.println("nDatiGrezzi = "+nDatiGrezzi);
		
		for(Row row: checksumsSheet) { //righe
			
			//System.out.println("row.getLastCellNum = "+row.getLastCellNum());
			
			for(int cn = 0; cn < row.getLastCellNum(); cn++) { //colonne
				
                if(indexRow > nDatiGrezzi && cn < 4) {
                	
                	Cell cell = row.getCell(cn ,MissingCellPolicy.CREATE_NULL_AS_BLANK) ;
                	cell.setCellType(CellType.BLANK);
                }
				
			}
			
			indexRow ++; 
		}
		
	}

/*--------------------------------------------------------[Fine]--------------------------------------------------------------*/
	
	

	
/* -----------------------------------------------------[checksums]----------------------------------------------------------*/
	/* Metodo che lavora il tab Checksums */
	 public void checksums(File f, File f3) {
	    	
	    	try {
	    		
	    		/* FileInputStream per il file in cui devo copiare i dati */
	    		FileInputStream fileVuoto = new FileInputStream( f ); //file vuoto
	    		
				/* Foglio Excel nuovo CM */
				XSSFWorkbook workbook_fv = new XSSFWorkbook(fileVuoto);
				
			    /*vado nel tab desiderato che in questo caso è Checksums */
				XSSFSheet tabChecksumsFileVuoto = workbook_fv.getSheetAt(1);
				
				/* prima di scrivere dentro il tab vuoto lo pulisco */
				deleteSheetAllContent(tabChecksumsFileVuoto);
				
				XSSFSheet tabDescriptionFileVuoto = workbook_fv.getSheetAt(7);
				
				/* carico il tab Description */
				caricaDescription(tabDescriptionFileVuoto, f3);
				
				/* carica gli ArrayList per column1 , 2 , 3 , 4 di checksums*/
				leggiVCM(f3);
				
	    		
				int nRigheTabVuoto = 0; 
				int indexPath = 1;
				int indexFileName = 2;
				int indexSha1 = 1;
				int indexDescr = 2;
				int indexCol1 = 1;
				int indexCol2 = 1;
				int indexCol3 = 1;
				int indexCol4 = 1;
				int indexCol5 = 2;
				int indexCol6 = 2;
				
				/* prima di scrivere dentro il tab vuoto lo pulisco */
				deleteSheetAllContent(tabChecksumsFileVuoto);
				
				Row titoloTabFileVuoto = tabChecksumsFileVuoto.createRow(nRigheTabVuoto); 
				titoloTabFileVuoto.createCell(0).setCellValue("FileName");
				titoloTabFileVuoto.createCell(1).setCellValue("Path");
				titoloTabFileVuoto.createCell(2).setCellValue("Sha1");
				titoloTabFileVuoto.createCell(3).setCellValue("Description");
				titoloTabFileVuoto.createCell(4).setCellValue("");
				titoloTabFileVuoto.createCell(5).setCellValue("Column1");
				titoloTabFileVuoto.createCell(6).setCellValue("Column2");
				titoloTabFileVuoto.createCell(7).setCellValue("Column3");
				titoloTabFileVuoto.createCell(8).setCellValue("Column4");
				titoloTabFileVuoto.createCell(9).setCellValue("For Lookup");
				titoloTabFileVuoto.createCell(10).setCellValue("Column5");
				
			
				nRigheTabVuoto ++; 
				
				//System.out.println("dim vettore1 = "+checksumColumn1.size());
				
				
				for(int i=0;i<checksumColumn1.size()-1; i++) {
					
				    Row rigaCreataFileVuoto = tabChecksumsFileVuoto.createRow(nRigheTabVuoto);
				    
				    String cellaDescription = "B"+indexDescr;
				    String matrice = "Description!B2:C139";
				    String formulaDescr = "VLOOKUP("+cellaDescription+","+matrice+",2,FALSE)";
				    
				    String cella = "B"+indexFileName;
				    String formula1 = "Grezzi!B"+indexPath;
		            String formula2 = "Grezzi!A"+indexSha1;
		            String formula3 = "RIGHT("+cella+",LEN("+cella+")-FIND("+cella+",SUBSTITUTE("+cella+","+"\"/\" "+","+cella+",LEN("+cella+")-LEN(SUBSTITUTE("+cella+","+"\"/\" "+","+""+"))),1))";
		            
		            String separator = "\"_\" ";
		            String formula4 = "CONCATENATE(F"+indexCol5+","+separator+","+"G"+indexCol5+")";
		            
		            String cellaColumn5 = "I"+indexCol6;
					String matriceColumn5 = "Checksums!B2:C122";
					String formula5 = "VLOOKUP("+cellaColumn5+","+matriceColumn5+",2,FALSE)";
		            
				    rigaCreataFileVuoto.createCell(1).setCellFormula(formula1);
				    rigaCreataFileVuoto.createCell(2).setCellFormula(formula2);
				    rigaCreataFileVuoto.createCell(0).setCellFormula(formula3);
				    rigaCreataFileVuoto.createCell(3).setCellFormula(formulaDescr);
				    rigaCreataFileVuoto.createCell(5).setCellValue(checksumColumn1.get(indexCol1));
				    rigaCreataFileVuoto.createCell(6).setCellValue(checksumColumn2.get(indexCol2));
				    rigaCreataFileVuoto.createCell(7).setCellValue(checksumColumn3.get(indexCol3));
				    rigaCreataFileVuoto.createCell(8).setCellValue(checksumColumn4.get(indexCol4));
				    nDatiChecksums++;
				    
				    rigaCreataFileVuoto.createCell(9).setCellFormula(formula4);
				    rigaCreataFileVuoto.createCell(10).setCellFormula(formula5);
				    
				    
				    nRigheTabVuoto ++;
				    indexPath ++;
				    indexFileName ++;
				    indexSha1 ++;
				    indexDescr ++;
				    
				    indexCol1++;
				    indexCol2++;
				    indexCol3++;
				    indexCol4++;
				    indexCol5++;
				    indexCol6++;
				    
				}
				
				deleteEmptyRows(tabChecksumsFileVuoto);
				
				
				//aggiorna il file vuoto
				FileOutputStream out = new FileOutputStream(f);
				workbook_fv.write(out);
				out.close();
				
				
	    	}catch(IOException ex) { ex.printStackTrace(); }

	    }

/*--------------------------------------------------------[Fine]-------------------------------------------------------------*/	

	 
	 
	 
	 
/*--------------------------------------------------------[CheckEVO]-----------------------------------------------------------*/
	 /* Metodo che lavora il tab Check EVO */
	 public void checkEVO(File nuovoCM , File fileCritical) {
		
		 try {
			 
	    	/* FileInputStream per il file del nuov CM */
	    	FileInputStream fileNuovoCM = new FileInputStream( nuovoCM ); //file nuovo CM
	    	
	    	/* FileInputStream per il file critical asset register  */
	    	FileInputStream fileCritAssetsRegister = new FileInputStream(fileCritical);
	    	
	    	/* foglio excel per il file critical assets register */
	    	XSSFWorkbook workbookCAR = new XSSFWorkbook(fileCritAssetsRegister);
	    	
	    	/* foglio excel per il file del nuovo CM */
	    	XSSFWorkbook w = new XSSFWorkbook(fileNuovoCM);
	    	
	    	
		    /*vado nel tab desiderato del file critical assets register  */
			XSSFSheet tabCAR = workbookCAR.getSheetAt(0);
	    	
	    	int numRowCAR = 0; 
	    	int numRowNuovoCM = 0; 
	    	
		    /*vado nel tab Check EVO nel file del nuovo CM  */
			XSSFSheet tabCheckEvoNuovoCM = w.getSheetAt(6);
	    	
			/* Pulisco il tab Check EVO e inserisco intanto il titolo */
			
			deleteSheetAllContent(tabCheckEvoNuovoCM);
			
			Row titoloTabNuovoCM = tabCheckEvoNuovoCM.createRow(numRowNuovoCM);
			titoloTabNuovoCM.createCell(0).setCellValue("Summary");
			titoloTabNuovoCM.createCell(1).setCellValue("Evo_Checksums");
			titoloTabNuovoCM.createCell(2).setCellValue("Path");
			titoloTabNuovoCM.createCell(3).setCellValue("Quinel_Filename");
			titoloTabNuovoCM.createCell(4).setCellValue("Q_Path");
			titoloTabNuovoCM.createCell(5).setCellValue("Compare");
			numRowNuovoCM ++; 
			
	        /* Iteratore righe file critical assets register */	
			Iterator<Row> carRowIterator = tabCAR.iterator() ; 
			
			/* Salto le prime 4 righe del file perché non mi interessano */
			for(int i=0;i<4;i++) {
				
				if(carRowIterator.hasNext()) {
					carRowIterator.next() ;
				}
			} 
			
			while(carRowIterator.hasNext()) {
				
				Row carRow = carRowIterator.next() ;
				
				Row nuovoCMRow = tabCheckEvoNuovoCM.createRow(numRowNuovoCM);
				
				int numCellCAR = 0; 
				Iterator<Cell> carCellIterator = carRow.cellIterator() ;
				
				while(carCellIterator.hasNext()) {
					
					Cell carCell = carCellIterator.next() ; 
					
					if(numCellCAR == 2) { 
						
						
						
						if( carCell.getStringCellValue().equals("certification-xml-daily") ) {
							
							nuovoCMRow.createCell(0).setCellValue(carCell.getStringCellValue());
							
						}else {
							
							/* Scrivo nome file */
							if(!carCell.getStringCellValue().endsWith(".class") ) {
								
								String str = carCell.getStringCellValue()+".class";
								
								if( str.contains("/") ) {
									
									int primoSlash = str.indexOf("/");
									String s = str.substring(primoSlash+1);
									nuovoCMRow.createCell(0).setCellValue(s);
									
								}else {
									
									nuovoCMRow.createCell(0).setCellValue(str);
								}
								
							}else if(carCell.getStringCellValue().endsWith(".class") ) {
								
								if( carCell.getStringCellValue().contains("/") ) {
									
									int primoSlash = carCell.getStringCellValue().indexOf("/");
									String s = carCell.getStringCellValue().substring(primoSlash+1);
									nuovoCMRow.createCell(0).setCellValue(s);
									
								}else {
									nuovoCMRow.createCell(0).setCellValue(carCell.getStringCellValue());
								}
								
							}
							
						}
						
						
						//nuovoCMRow.createCell(0).setCellValue(carCell.getStringCellValue());
						
						/*ne approfitto per creare le altre celle cosi dopo devo solo cambiare il valore */
						nuovoCMRow.createCell(3).setCellValue("ciao");
						nuovoCMRow.createCell(4).setCellValue("ciao");
						nuovoCMRow.createCell(5).setCellValue("ciao");
						
					}else if(numCellCAR == 10) {
						
						/* Scrivo nome sha1 */
						nuovoCMRow.createCell(1).setCellValue(carCell.getStringCellValue());
						
					}else if(numCellCAR == 11) {
						
						/* Scrivo nome path */
						nuovoCMRow.createCell(2).setCellValue(carCell.getStringCellValue());
					}
					
					numCellCAR ++; 
				}
			    
				//System.out.println("\n");
				numRowNuovoCM ++;
				numRowCAR ++; 
				
			}
			
			/*Doper aver riempito le prime tre colonne devo riempire le ultime tre */
			int nRow = 0; 
			int indexColoumnB = 2;
			int indexColoumnE = 2;
			int indexColumnF = 2; 
			
			for(Row row: tabCheckEvoNuovoCM) {
				
				for(int cn=0; cn < row.getLastCellNum(); cn ++) {
					
					if(cn == 3 && nRow > 0) {
						
						String cellaColumnB = "B"+indexColoumnB;
						String cellaColumnE = "E"+indexColoumnE;
						
						String matrice = "Grezzi!A:B";
						String formula = "VLOOKUP("+cellaColumnB+","+matrice+",2,FALSE)";
						String formulaEstrai ="RIGHT("+cellaColumnE+",LEN("+cellaColumnE+")-FIND("+cellaColumnE+",SUBSTITUTE("+cellaColumnE+","+"\"/\" "+","+cellaColumnE+",LEN("+cellaColumnE+")-LEN(SUBSTITUTE("+cellaColumnE+","+"\"/\" "+","+""+"))),1))";
						String formulaEquals= "IF(A"+indexColumnF+"=D"+indexColumnF+",TRUE)"; 
						
						row.getCell(4).setCellFormula(formula);
						row.getCell(3).setCellFormula(formulaEstrai);
						row.getCell(5).setCellFormula(formulaEquals);
						
						indexColoumnB ++;
						indexColoumnE ++;
						indexColumnF ++;
					}
				}
				
				nRow ++;
			}
			
			//System.out.println("numRowCAR : "+numRowCAR);
			
			//aggiorna il file del nuovo CM 
			FileOutputStream out = new FileOutputStream(nuovoCM);
			w.write(out);
			out.close();
	    	
		 }catch(IOException ex) { ex.printStackTrace(); }
	 }
	 
/*----------------------------------------------------------[Fine]-------------------------------------------------------------*/	 

	 
	 

/*-----------------------------------------------------[leggiEvoGameVersions]----------------------------------------------------*/	 
	 
	 
/*-------------------------------------------------------------[Fine]------------------------------------------------------------*/	 
	 
	 
	 public void leggiEvoGameVersions(XSSFSheet mySheet) {
		 
		 int rowNum = 0; 
		 
		 for(Row row: mySheet) {
			 
			 for(int cn = 0; cn < row.getLastCellNum(); cn++) {
				 
				 Cell cell = row.getCell(cn ,MissingCellPolicy.CREATE_NULL_AS_BLANK) ;
				 
				 if( cn == 0 && rowNum > 0) {
					 
		    			if(cell.toString().equals("Blackjack RNG")) {
		    				
		    				EGVGameName.add("First Person Blackjack");
		    				
		    			}else if(cell.toString().equals("Roulette RNG")) {
		    				
		    				EGVGameName.add("First Person Roulette");
		    				
		    			}else {
		    				
		    				EGVGameName.add(cell.toString());
		    			}
					
					 
				 }else if( cn == 1 && rowNum > 0) {
					 
					 EGVGameType.add(cell.toString());
					 
				 }else if( cn == 2 && rowNum > 0) {
					 
					EGVPlatformVersion.add(cell.toString());
					 
				 }else if( cn == 3 && rowNum > 0) {
					 
					 EGVGameVersion.add(cell.toString()) ;
				 }
				 
			 }
			 rowNum ++;
		 }
	 }
	 
	 
	 
/*------------------------------------------------------------[Report]----------------------------------------------------------*/
	 /* Metodo che lavora il tab Report */
	 public void report(File nuovoCM, File gameVersions) {
		 
		 try {
			 
		    	/* FileInputStream per il file del nuovo CM */
		    	FileInputStream fileNuovoCM = new FileInputStream( nuovoCM );
		    	
		    	/* FileInputStream per il file gameVersions  */
		    	FileInputStream fileGameVersions = new FileInputStream(gameVersions);
		    	
		    	/* foglio excel per il file Evolution Game Versions */
		    	XSSFWorkbook workbook_gv = new XSSFWorkbook(fileGameVersions);
		    	
		    	/* foglio excel per il file del nuovo CM */
		    	XSSFWorkbook workbook_ncm = new XSSFWorkbook(fileNuovoCM);
		    	
		    	
		    	/*Tab Version del file Evolution Game Versions*/
		    	XSSFSheet tabFileGameVersions = workbook_gv.getSheetAt(0);
		    	
		    	/*Tab Report del file del nuovo CM */
		    	XSSFSheet tabReportFileNuovoCM = workbook_ncm.getSheetAt(2);
		    	
		    	/*Chiamo la funzione che mi legge il contenuto del tab version del file Evolution Games versions e
		    	 *  va a memorizzare i dati dentro degli ArrayList */
		    	leggiEvoGameVersions(tabFileGameVersions);
		    	
		    	/* Contatore righe tabella grande */
		    	int nRowBigTable = 0; 
		    	
		    	/* Contatore righe tabella piccola */
		    	int nRowSmallTable = 5;
		    	
		    	/*Indice degli ArrayList per il riempito della tabella piccola*/
		    	int n = 0; 
		    	
		    	int nsTable = 0;
		    	
		    	/* Aggiungo il titolo della tabella grande */
		    	Row titleBigTable = tabReportFileNuovoCM.createRow(nRowBigTable);
		    	titleBigTable.createCell(0).setCellValue("Game");
		    	titleBigTable.createCell(1).setCellValue("FileName");
		    	titleBigTable.createCell(2).setCellValue("Sha1");
		    	titleBigTable.createCell(3).setCellValue("Platform Version");
		    	titleBigTable.createCell(4).setCellValue("Game Version");
		    	titleBigTable.createCell(8).setCellValue("RNG + Comuni");
		    	
		    	nRowBigTable ++;
		    	
		    	/*Inserimento dei dati nella tabella grande , Sono quelli che ho già letto nel tab Checksums 
		    	 * del file del nuovo CM e memorizzato dentro i vettori checksumsColumn1 , checksumsColumn2 , 
		    	 * checksumsColumn5 */
		    	
		    	for(int i=0; i<checksumsColumn1.size(); i++) {
		    		
		    		Row contentsBigTable = tabReportFileNuovoCM.createRow(nRowBigTable);
		    		
		    		contentsBigTable.createCell(0).setCellValue(checksumsColumn1.get(n));
		    		contentsBigTable.createCell(1).setCellValue(checksumsColumn2.get(n));
		    		contentsBigTable.createCell(2).setCellValue(checksumsColumn5.get(n));
		    		contentsBigTable.createCell(3).setCellValue("empty");
		    		contentsBigTable.createCell(4).setCellValue("empty");
		 		    		
		    		
		    		/*Devo intanto proseguire con l'inserimento del titolo della tabella piccola*/
		    		if(nRowBigTable == 5) { 
		    			
				    	/* Aggiungo il titolo della tabella piccola ma saltando 4 righe */
		    			contentsBigTable.createCell(8).setCellValue("GameName");
		    			contentsBigTable.createCell(9).setCellValue("GameType");
		    			contentsBigTable.createCell(10).setCellValue("Platform Version");
		    			contentsBigTable.createCell(11).setCellValue("Game Version");
				    	nRowSmallTable ++;
				    	
		    		}else if( nRowBigTable > 5 && nsTable < EGVGameName.size()) {
		    			
		    			/* Inserimento dei dati nella piccola tabella 
				    	 * Sono quelli che ho letto prima chiamando la funzione leggiEvoGameVersions*/
		    			
		    			contentsBigTable.createCell(8).setCellValue(EGVGameName.get(nsTable));
		    			contentsBigTable.createCell(9).setCellValue(EGVGameType.get(nsTable));
		    			contentsBigTable.createCell(10).setCellValue(EGVPlatformVersion.get(nsTable));
		    			contentsBigTable.createCell(11).setCellValue(EGVGameVersion.get(nsTable));
		    			
		    			nsTable ++;
		    			
		    		}else if( nRowBigTable == 1 ) {
		    			
		    			contentsBigTable.createCell(8).setCellValue("RngService.class");
		    			contentsBigTable.createCell(9).setCellValue(RngService);
		    			
		    		}else if( nRowBigTable == 2 ) {
		    			
		    			contentsBigTable.createCell(8).setCellValue("prepare_checksums.py");
		    			contentsBigTable.createCell(9).setCellValue(Sha1PrepareChecksumsGrezzi);
		    		}
		    		
		    		n++;
		    		nRowBigTable ++;
		    	}
		    	
		    	
		    	/* Ora devo riempire la colonna Platform Version , Game Version , e Sì della tabella grande 
		    	 * perché ora ho tutti i dati che mi servono a disposizione */
		    	
		    	int rowNum = 0; 
		    	int indexCellToSearch = 2 ; 
		    	
		    	for(Row row: tabReportFileNuovoCM) {
		    		
		    		for(int cn=0; cn < row.getLastCellNum(); cn++) {
		    			
		    			Cell cell = row.getCell(cn ,MissingCellPolicy.CREATE_NULL_AS_BLANK) ;
		    			
		    			if(cn == 3 && rowNum > 0) {
		    				
		    				String cellToSearch = "A"+indexCellToSearch;
		    				String matrice = "Report!I5:L46";
		    				
							String formulaPlatversion = "VLOOKUP("+cellToSearch+","+matrice+",3,FALSE)";
							cell.setCellFormula(formulaPlatversion);
							
							
		    			}else if(cn == 4 && rowNum > 0) {
		    				
		    				String cellToSearch = "A"+indexCellToSearch;
		    				String matrice = "Report!I5:L46";
		    				
							String gameVersion = "VLOOKUP("+cellToSearch+","+matrice+",4,FALSE)";
							cell.setCellFormula(gameVersion);
							
							indexCellToSearch ++;
		    			}
		    			
		    		}
		    		
		    		rowNum ++;
		    	}
		    	
		    	
				//aggiorna il file del nuovo CM 
				FileOutputStream out = new FileOutputStream(nuovoCM);
				workbook_ncm.write(out);
				out.close();
		    	
		 }catch( IOException ex) { ex.printStackTrace();}
	 }
	 
/*-------------------------------------------------------------[Fine]------------------------------------------------------------*/


	 
	 
/*----------------------------------------------------------[leggiGameVersionsVCM]------------------------------------------------*/
	 
	 /* Metodo che legge il contenuto del tab Game Versions del file del vecchio CM e memorizza il GameName */
	 
	 public void leggiGameVersionsVCM(XSSFSheet mySheet) {
		 
		 int rowNum = 0; 
		 
		 for(Row row: mySheet) {
			 
			 for(int cn=0; cn<row.getLastCellNum(); cn++) {
				 
				 Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
						 
				 if( cn == 0 && rowNum > 0) {
					 gameVersionsGameName.add(cell.toString()) ; 
				 }
			 }
			 
			 rowNum ++;
		 }
	 }
	 
/*------------------------------------------------------------------[Fine]-------------------------------------------------------*/	 
	 
	 
	 
	 
/*------------------------------------------------------------[GameVersions]-------------------------------------------------------*/
	 /* Metodo che lavora dentro il tab Game Versions */
	 public void gameVersions(File nuovoCM , File vecchioCM) {
		 
		 try {
			 
		    	/* FileInputStream per il file del nuovo CM */
		    	FileInputStream fileNuovoCM = new FileInputStream( nuovoCM );
		    	
		    	/* FileInputStream per il file del vecchio CM */
		    	FileInputStream fileVecchioCM = new FileInputStream( vecchioCM );
		    	
		    	/* foglio excel per il file del nuovo CM */
		    	XSSFWorkbook workbook_ncm = new XSSFWorkbook(fileNuovoCM);
		    	
		    	/* foglio excel per il file del vecchio CM */
		    	XSSFWorkbook workbook_vcm = new XSSFWorkbook(fileVecchioCM);
		    	
		    	/*Tab Game Versions del file del nuovo CM */
		    	XSSFSheet tabGameVersionsFileNuovoCM = workbook_ncm.getSheetAt(3);
		    	
		    	/*Tab Game Versions del file del vecchio CM */
		    	XSSFSheet tabGameVersionsFileVecchioCM = workbook_vcm.getSheetAt(3);
		    	
		    	/* Leggo e memorizzo la prima colonna (Game Name) del tab Game Versions del vecchio CM */
		    	leggiGameVersionsVCM(tabGameVersionsFileVecchioCM);
		    	
		    	/*Contatore righe tab Game Versions File nuovo CM */
		    	int rowNumTabGameVersions = 0;
		    	int indexValue = 0; 
		    	
		    	/*Inserisco il titolo */
		    	Row header = tabGameVersionsFileNuovoCM.createRow(rowNumTabGameVersions); 
		    	header.createCell(0).setCellValue("GameName");
		    	header.createCell(1).setCellValue("Platform Version");
		    	header.createCell(2).setCellValue("Game Version");
		    	header.createCell(3).setCellValue("Note");
		    	
		    	rowNumTabGameVersions ++;
		    	
		    	/*Ora inserisco il contenuto*/
		    	
		    	int indexCellToSearch = 2;
		    	
		    	for(int i=0; i<gameVersionsGameName.size(); i++) {
		    		
		    		Row contentsRow = tabGameVersionsFileNuovoCM.createRow(rowNumTabGameVersions);
		    		contentsRow.createCell(0).setCellValue(gameVersionsGameName.get(indexValue));
		    		
    				String cellToSearch = "A"+indexCellToSearch;
    				String matrice = "Report!I5:L46";
    				
					String formulaPlatversion = "VLOOKUP("+cellToSearch+","+matrice+",3,FALSE)";
					String formulaGameVersion = "VLOOKUP("+cellToSearch+","+matrice+",4,FALSE)";
					
					contentsRow.createCell(1).setCellFormula(formulaPlatversion);
					contentsRow.createCell(2).setCellFormula(formulaGameVersion);
					
					if( gameVersionsGameName.get(indexValue).equals("2 Hand Casino Hold’em") ) {
						
						contentsRow.createCell(3).setCellValue("Changed name");
						
					}else if( gameVersionsGameName.get(indexValue).equals("Baccarat") ) {
						
						contentsRow.createCell(3).setCellValue("Removed(*)");
						
					}else if( gameVersionsGameName.get(indexValue).equals("Blackjack") ) {
						
						contentsRow.createCell(3).setCellValue("Removed(**)");
						
					}else if( gameVersionsGameName.get(indexValue).equals("Roulette") ) {
						
						contentsRow.createCell(3).setCellValue("Removed(***)");
					}
					
					indexCellToSearch ++;
		    		indexValue ++;
		    		rowNumTabGameVersions ++;
		    		
		    	}
		    	
		    	
				//aggiorna il file del nuovo CM 
				FileOutputStream out = new FileOutputStream(nuovoCM);
				workbook_ncm.write(out);
				out.close();
		    	
		 }catch(IOException ex) { ex.printStackTrace(); }
		 
	 }
	 
	 
/*---------------------------------------------------------------[Fine]------------------------------------------------------------*/
	 


/*-----------------------------------------------------------[ChangedGames]-------------------------------------------------------*/	 
	 /*Metodo che lavora dentro il tab Changed Games*/
	 public void changedGames(File nuovoCM) {
		 
		 try {
			 
		    	/* FileInputStream per il file del nuovo CM */
		    	FileInputStream fileNuovoCM = new FileInputStream( nuovoCM );
		    	
		    	/* foglio excel per il file del nuovo CM */
		    	XSSFWorkbook workbook_ncm = new XSSFWorkbook( fileNuovoCM );
		    	
		    	/*Tab Changed Games del file del nuovo CM */
		    	XSSFSheet tabChangedGamesNuovoCM = workbook_ncm.getSheetAt(4);
		    	
		    	int rowNum = 0;
		    	int index = 0; 
		    			
		    	/*Inserimento del titolo tab Changed Games file nuovo CM */
		    	Row header = tabChangedGamesNuovoCM.createRow(rowNum); 
		    	header.createCell(0).setCellValue("GameName");
		    	header.createCell(1).setCellValue("Changed");
		    	rowNum ++;
		    	
		    	/* Ora inserisco la prima colonna che sono i GameName */
		    	for(int i=0; i<gameVersionsGameName.size(); i++) {
		    		
		    		Row contentRow = tabChangedGamesNuovoCM.createRow(rowNum);
		    		contentRow.createCell(0).setCellValue(gameVersionsGameName.get(index));
		    		contentRow.createCell(1).setCellValue("empty");
		    		index ++;
		    		rowNum ++;
		    	}
		    	
		    	
		    	/*Ora controllo i giochi che sono cambiati , praticamente devo confrontare ogni 
		    	 * C_GAME con O_GAME, C_FILE con O_FILE , C_SHA1 con O_SHA1 come avevo già fatto
		    	 * in Appoggio Changed Games */
		    	
		    	
				//aggiorna il file del nuovo CM 
				FileOutputStream out = new FileOutputStream(nuovoCM);
				workbook_ncm.write(out);
				out.close();
		    	
			 
		 }catch(IOException ex) { ex.printStackTrace(); }
		 
	 }
/*---------------------------------------------------------------[Fine]------------------------------------------------------------*/
	 
}
