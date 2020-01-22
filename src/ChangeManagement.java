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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe che gestisce tutto il Change Management. 
 * Ciascun metodo ha un ruolo ben preciso , per esempio:
 * il metodo "appoggioChangedGames()" andra nel tab "Appoggio Changed Games" 
 * e gestira tutto il lavoro che dovr� fare in "Appoggio Changed Games" e restituir�
 * il file aggiornato,
 * il metodo "checkSums()" andr� nel tab "Checksums" e gestira tutto il lavoro che dovr�
 * fare in checksum e restituira il file aggiornato e cos� via.
 * Author : Dedou Maximin */

public class ChangeManagement {

	/* Il foglio excel che contiene tutto il Change Management */
	private XSSFWorkbook workbook;
	
	/* I vari tabs del foglio excel */
	private XSSFSheet [] tabs;
	
	private String nomeFoglioExcel;
	
	/* controllare se la creazione del foglio è andata buon fine */
	private int rt ; 

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
		
		//creo fisicamente il foglio excel
		rt = generaFoglioExcel(nomeFoglioExcel);
		
	}
	
	/* Metodo che ritorna rt*/
    public int getRt() {
    	return rt; 
    }
    
	/* Metodo che restituisce il foglio Excel*/
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}
    
	/* Metodo che modifica il foglio excel*/
	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}
	
	
	public XSSFSheet[] getTabs() {
		return tabs;
	}

	public void setTabs(XSSFSheet[] tabs) {
		this.tabs = tabs;
	}
	
	
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
	
	
	/* metodo che mi sposta la colonna la colonna degli di appoggio changed games dalla 3� colonna alla 6� */
	public void spostaSha1AppChGames(XSSFSheet mySheet) {
		
		Iterator<Row> RowIterator = mySheet.iterator();
		int rowNum = 0; 
		
		while(RowIterator.hasNext()) {
			
			int cellNum = 0; 
			Row row = RowIterator.next(); 
			Iterator<Cell> cellIterator = row.cellIterator();
			
			String C_Sha1 = "";
			
			while(cellIterator.hasNext()) {
				 
				
				Cell cell = cellIterator.next(); //prendo ogni cella 
				
				if(cellNum == 0) {
					C_Sha1 = cell.getStringCellValue() ; 
					cell.setCellType(CellType.BLANK);
					//System.out.println(C_Sha1);
				}else if(cellNum == 1) {

				}else if(cellNum == 3){
					cell.setCellValue(C_Sha1);
				}
				//System.out.println("\n");
				cellNum ++;
			}
			
			rowNum ++; 
		}
	}

	/* Metodo che lavora all'interno del tab Appoggio Changed Games 
	 * getRichStringCellValue().toString();*/
	public void appoggioChangedGames(File f , File f2) {
		
		try {
			
			FileInputStream fileVuoto = new FileInputStream( f ); //file vuoto
			FileInputStream filePieno = new FileInputStream(f2) ; // file pieno 
			
			XSSFWorkbook workbook_fv = new XSSFWorkbook(fileVuoto);
			XSSFWorkbook workbook_fp = new XSSFWorkbook(filePieno) ; 
			
		    //vado nel tab desiderato che in questo caso � appoggio changed games
			XSSFSheet desiredSheetV = workbook_fv.getSheetAt(5);
			XSSFSheet desiredSheetP = workbook_fp.getSheetAt(5);
			
			int rowNum = 0; 
			Row header = desiredSheetV.createRow(rowNum);
			
		    header.createCell(0).setCellValue("C_Game");
		    header.createCell(1).setCellValue("C_File");
		    header.createCell(2).setCellValue("C_Sha1");
		    header.createCell(3).setCellValue("O_Game");
		    header.createCell(4).setCellValue("O_File");
		    header.createCell(5).setCellValue("O_Sha1");
		    header.createCell(6).setCellValue("Game");
		    header.createCell(7).setCellValue("File");
		    header.createCell(8).setCellValue("Sha");
			
			
			Iterator<Row> RowIterator = desiredSheetP.iterator();
			
			//se � la prima riga inserisco il titolo
			if( RowIterator.hasNext() ) {
				RowIterator.next();
				rowNum ++; 
			}
			
			int indexRow1 = 2;
			int indexRow2 = 2; 
			int indexRow3 = 2; 
			
			while(RowIterator.hasNext()) {
				
				Row row = RowIterator.next(); 
				
				Row row2 = desiredSheetV.createRow(rowNum) ; 
				
				Iterator<Cell> cellIterator = row.cellIterator();
				
				int cellNum = 0;
				
				String C_Game = ""; 
				String C_File = ""; 
				String C_Sha1 = ""; 
				
				if(cellIterator.hasNext()) {
					//creo una cella nell'altro foglio
					Cell c = row2.createCell(cellNum);
					
					//invece nel foglio corrente memorizzo i dati
					Cell cell = cellIterator.next() ; 
					C_Game = cell.getStringCellValue() ; 		
					
					cellNum ++; 
				}
				
				if(cellIterator.hasNext()) {
					//creo una cella nell'altro foglio
					Cell c = row2.createCell(cellNum);
					
					//invece nel foglio corrente memorizzo i dati
					Cell cell = cellIterator.next() ; 
					C_File = cell.getStringCellValue() ; 		
					
					cellNum ++; 
				}
				
				if(cellIterator.hasNext()) {
					//creo una cella nell'altro foglio
					Cell c = row2.createCell(cellNum);
					
					//invece nel foglio corrente memorizzo i dati
					Cell cell = cellIterator.next() ; 
					
					C_Sha1 = cell.getStringCellValue() ;
					
					cellNum ++; 
				}
				
				while(cellIterator.hasNext()) {
					 
					Cell cell = cellIterator.next(); //prendo ogni cella 
					
					switch(cell.getCellType()) {
					
					  case NUMERIC: 
						  
						break;
					  
					  case STRING: 
						  
						  if(cellNum == 3) { //C_GAME
							  Cell c = row2.createCell(cellNum); 
							  c.setCellValue(C_Game);
						  }else if(cellNum == 4) { //C_FILE
							  Cell c = row2.createCell(cellNum) ; 
							  c.setCellValue(C_File);
							  
						  }else if(cellNum == 5) {
							  Cell c = row2.createCell(cellNum) ; 
							  c.setCellValue(C_Sha1);
						  }else {
							  Cell c = row2.createCell(cellNum) ; 
						  }
						  
						  cellNum++; 
						  
						  
						break;	
						
					  case FORMULA:
						
						 
						 if(cellNum == 2) {
							 
							 switch(cell.getCachedFormulaResultType()) {
							  
							  case STRING:  
								  Cell c = row2.createCell(cellNum) ;
								  c.setCellValue(cell.getRichStringCellValue().toString());
							  break ; 

							  case BOOLEAN:
								  Cell c2 = row2.createCell(cellNum);
								  break;
						     }
							 
						 }else if(cellNum == 6 || cellNum == 7 || cellNum == 8){
							 
							 String strFormula = ""; 
							 
							 if(cellNum == 6) {
								 strFormula = "IF(A"+indexRow1+"=D"+indexRow1+",TRUE)";
								 indexRow1 ++;
								 
							 }else if(cellNum == 7) {
								 strFormula = "IF(B"+indexRow2+"=E"+indexRow2+",TRUE)";
								 indexRow2 ++;
								 
							 }else if(cellNum == 8) {
								 strFormula = "IF(C"+indexRow3+"=F"+indexRow3+",TRUE)";
								 indexRow3 ++;
							 }
							
							 Cell c = row2.createCell(cellNum);
							 c.setCellFormula(strFormula);
							 
						 }

						 cellNum++; 
						  
						  break; 
					}
					
				}
				
				cellNum = 0; 
				
				System.out.println("\n");
				
				
				rowNum ++;
			}
			
			
			filePieno.close(); 
			fileVuoto.close();
			
			//aggiorna il file che era vuoto
			FileOutputStream out = new FileOutputStream(f);
			workbook_fv.write(out);
			out.close();

		}catch(Exception ex) { ex.printStackTrace(); }
	}
	
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
	
	
	
	/* Metodo che cancella tutto il contenuto di una certa colonna dentro un preciso tab */
	public void deleteContentsAddFormulaColumn(XSSFSheet desiredSheet , int nColumn) throws IOException {
		
		Iterator<Row> rowIterator = desiredSheet.iterator();
		
		int rowNum = 0 ; 
		/* Se � la prima riga la salto perch� titolo */
		if( rowIterator.hasNext() ) {
			rowIterator.next();
			rowNum ++;
		}
		
        while (rowIterator.hasNext()) 
        {
            Row row = rowIterator.next();
            //For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();
            int cellNum = 0; 
            
            while (cellIterator.hasNext()) 
            {
                Cell cell = cellIterator.next();
                
                if( nColumn >= 0 ) {
                	
                	if(cellNum == nColumn) {
                		
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
                            	
                            	//aggiungo la formula dopo aver cancellato
                            	cell.setCellFormula("Grezzi!B"+(rowNum) );
                            	//System.out.println("rowNum = "+rowNum);
                                break;    
                        }
                        
                	}
                }else {
                	System.out.println("indice colonna sbagliata...cancellazione impossibile!");
                }
                
                cellNum ++ ;
                 
            }
            
            rowNum ++;
            
        }
        
	}
	
	
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
			
			//vado nel tab desiderato che in questo caso � Grezzi 
			XSSFSheet desiredSheet = workbook.getSheetAt(0);
			System.out.print("Aggiornamento tab: "+desiredSheet.getSheetName()+"...");
			
			/* cancello tutti i dati che ci sono gi� nel tab grezzi */
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
			
	        while( (st = br.readLine()) != null ) {
	        	appoggio = sostituisci(st);
	        	String dati = new String(appoggio);
	        	String sha1 = dati.substring(0,40);
	        	String path = dati.substring(41);
	        	
	        	Row row = desiredSheet.createRow(rowNum) ;
	        	
	        	int cellNum = 0; 
	        	for(int i=0;i<2; i++) {
	        		
	        		Cell cell = row.createCell(i);
	        		if(i == 0) {
	        			cell.setCellValue(sha1);
	        		}else {
	        			
	        			/* tolgo dal path il nome della cartella principale del gioco */
	        			int posPrimoSlash = path.indexOf("/") ;
	        			String nuovoPath = path.substring(posPrimoSlash+1); 	        			
	        			cell.setCellValue(nuovoPath);
	        		}
	        	}
	        	 
	        	rowNum ++; 
	        }
	        
	        /* add file prepare_checksums.py*/
	        Row r = desiredSheet.createRow(rowNum++);
	        for(int i=0;i<2;i++) {
	        	Cell c = r.createCell(i);
        		if(i == 0) {
        			c.setCellValue(sha1CheckSum.toLowerCase());
        		}else {
        			c.setCellValue(nomeCheckSum);
        		}
	        }
	        
	        
			//aggiorna il foglio
			FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			System.out.print(" Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
		}else {
			System.out.println("File non trovato! ");
		}
		
	}
	
	
	
	/* Metodo che lavora dentro il tab Checksums */
	/*
	 * Formula parola dopo ultimo slash(col A): "STRINGA.ESTRAI([@Path];TROVA("*";SOSTITUISCI([@Path];"/";"*";LUNGHEZZA([@Path])-LUNGHEZZA(SOSTITUISCI([@Path];"/";""))))+1;LUNGHEZZA([@Path]))" 
	 * Formula copia colonna tab(col B): "=Grezzi!B1" 
	 * Formula descrizione(col D): "=CERCA.VERT([@Path];Summary6[[Path]:[Description]];2;FALSO)" 
	 * Formula (col C): "=Grezzi!A1" 
	 * Formula (col J): "=[@Column1]&"_"&[@Column2]"
	 * Formula (col K): "=CERCA.VERT(Checksums!$I2;Summary[[Path]:[Sha1]];2;FALSO)"
	 * */
	public void checksums(File f) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( f ); //file nuovo CM
			
			XSSFWorkbook workbook_fp = new XSSFWorkbook(mioFile);
			
			XSSFSheet checksumsSheet = workbook_fp.getSheetAt(1);
			XSSFSheet grezziSheet = workbook_fp.getSheetAt(0);
			
			int rowNum = 0;
			int checksumsRowNum = 0; 
			
			/*inserimento titolo*/
			Row header = checksumsSheet.createRow(rowNum);
			
		    header.createCell(0).setCellValue("FileName");
		    header.createCell(1).setCellValue("Path");
		    header.createCell(2).setCellValue("Sha1");
		    header.createCell(3).setCellValue("Column1");
		    header.createCell(4).setCellValue("");
		    header.createCell(5).setCellValue("Column1");
		    header.createCell(6).setCellValue("Column2");
		    header.createCell(7).setCellValue("Column3");
		    header.createCell(8).setCellValue("Column4");
		    header.createCell(9).setCellValue("For Lookup");
		    header.createCell(10).setCellValue("Column5");
		    
		    
			
			/* ora scorro tutto Grezzi e per ogni riga in grezzi faccio la formula*/
		    Iterator<Row> rowIterator = grezziSheet.rowIterator(); 
		    
		    //salto la prima riga nel tab Checksums perch� � il titolo
		    Iterator<Row> checksumsRowIterator = checksumsSheet.rowIterator();
		    if(checksumsRowIterator.hasNext()) {
		    	checksumsRowIterator.next() ; 
		    	checksumsRowNum ++; 
		    }
	 
		    while(rowIterator.hasNext()) {
		    	
		    	Row row = rowIterator.next() ; 
		    	
		    	Row rowChecksums = checksumsSheet.createRow(checksumsRowNum) ; 
		    			
		    	Iterator<Cell> cellIterator = row.cellIterator() ; 
		    	
		    	int cellNum = 0; 
		    	//int checksumsCellNum = 0 ; 
		    	
		    	String sha1 = ""; 
		    	String path = ""; 
		    	
		    	//creo 11 celle in checksums per ogni riga creata
		    	for(int i=0; i<11;i++) {
		    		 rowChecksums.createCell(i);  
		    	}
		    	
		    	while(cellIterator.hasNext()) {
		    		
		    		Cell cell = cellIterator.next() ;
		    	    
		    		
		    		switch(cell.getCellType()) {
		    		
			    		case NUMERIC:
			    			break ; 
			    		case STRING:
			    			
			    			if(cellNum == 0) { //ho gli sha1
			    				
			    				sha1 = cell.getStringCellValue() ; 
			    				Cell cellChecksumsUpdate = checksumsSheet.getRow(checksumsRowNum).getCell(2);
			    				cellChecksumsUpdate.setCellValue(sha1);
			    				
			    			}else if(cellNum == 1) { //ho i path
			    				
			    			    path = cell.getStringCellValue() ; 	
			    			}
			    			break ; 
		    		}
		    		
		    		cellNum ++ ; 
		    
		    	}
		    	
		    	rowNum ++ ; 
		    	checksumsRowNum ++ ; 
		    }
		    
		    
			mioFile.close();
			
			//aggiorna il file che era vuoto
			FileOutputStream out = new FileOutputStream(f);
			workbook_fp.write(out);
			out.close();
		
			
		}catch(IOException ex) { ex.printStackTrace(); }
		
	}
	
}
