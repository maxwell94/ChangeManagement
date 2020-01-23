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
	
	ArrayList<String> checksumColumn1 ;
	ArrayList<String> checksumColumn2 ;
	ArrayList<String> checksumColumn3 ;
	ArrayList<String> checksumColumn4 ;

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
	
	
	/* metodo che mi sposta la colonna la colonna degli di appoggio changed games dalla 3° colonna alla 6° */
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
			
		    //vado nel tab desiderato che in questo caso è appoggio changed games
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
			
			//se è la prima riga inserisco il titolo
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
		/* Se è la prima riga la salto perché titolo */
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
	        			//System.out.println("path: "+path);
	        			int posPrimoSlash = path.indexOf("/");
	        			String nuovoPath = path.substring(posPrimoSlash+1); 	        			
	        			cell.setCellValue(nuovoPath);
	        			giochiPathGrezzi.add(nuovoPath); 
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
			//System.out.print(" Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
		}else {
			System.out.println("File non trovato! ");
		}
		
	}
	
	
	public int cercaPaths(String s) {
		
		int trovato = 0; 
		for(String str: giochiPathGrezzi) {
			if(s.equals(str)) {
				trovato = 1 ; 
			}
		}
		
		return trovato ; 
	}
	
	
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
							
							if(cellNum == 1) {
							   
								
							  int trovato = cercaPaths( cell.getStringCellValue() );
							  if( trovato == 0) {
								  //c.setCellValue("non trovato");
								  int primoSlash = cell.getStringCellValue().indexOf("/");
								  String newPth = cell.getStringCellValue().substring(primoSlash+1) ;
								  trovato = cercaPaths(newPth);
								  if(trovato == 1) {
									  c.setCellValue(newPth);
								  }else {
									  c.setCellValue(cell.getStringCellValue());
									  //System.out.println("controllare bene il path : "+newPth);
								  }
								  //c.setCellValue(newPth);
								  
							  }else if( trovato == 1) {
								  c.setCellValue(cell.getStringCellValue());
							  }
								
							}else {
								
								c.setCellValue(cell.getStringCellValue());
							}
							
							break ; 
					}
					
					cellNum ++; 
				}
				
				rowNum ++; 
			}
			
			
		}catch(IOException ex) {ex.printStackTrace();}
		 
	}
	
	
	
	/* Metodo che va a leggere il foglio excel del vecchio CM e memorizza la column1 , column2 , column3 , column4 dentro degli arraylist */
	public void leggiVCM(File f) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( f );	
			XSSFWorkbook w = new XSSFWorkbook(mioFile);
			
			//FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator() ; 
			
			XSSFSheet wSheet = w.getSheetAt(1); //vado nel tab Checksums
			
			Iterator<Row> rowIterator = wSheet.iterator() ; 
			
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
						checksumColumn4.add(cell.toString());
						//System.out.print(cell.toString()) ;
					}
					
				}
				
				System.out.println("\n");
			}
			
			
		}catch(IOException ex) { ex.printStackTrace(); }
	
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
	public void checksums1(File f, File v) {
		
		try {
			
			FileInputStream mioFile = new FileInputStream( f ); //file nuovo CM
			
			//FileInputStream fileVCM = new FileInputStream( v ); //file vecchio CM
			
			XSSFWorkbook workbook_fp = new XSSFWorkbook(mioFile);
			
			XSSFSheet checksumsSheet = workbook_fp.getSheetAt(1);
			XSSFSheet grezziSheet = workbook_fp.getSheetAt(0);
			
			XSSFSheet descriptionSheet = workbook_fp.getSheetAt(7);
			
			
			int rowNum = 0;
			int checksumsRowNum = 0; 
			int checksumsSheetVcmRowNum = 0; 
			
			/*inserimento titolo nuovo foglio Excel */
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
		    
		    //salto la prima riga nel tab Checksums perché è il titolo
		    Iterator<Row> checksumsRowIterator = checksumsSheet.rowIterator();
		    if(checksumsRowIterator.hasNext()) {
		    	checksumsRowIterator.next() ; 
		    	checksumsRowNum ++; 
		    }
		    
		    int indexRow = 1; 
	 
		    while(rowIterator.hasNext()) {
		    	
		    	Row row = rowIterator.next() ; 
		    	

		    	
		    	Row rowChecksums = checksumsSheet.createRow(checksumsRowNum) ; 
		    			
		    	Iterator<Cell> cellIterator = row.cellIterator() ; 
		    	
		    	int cellNum = 0; 
		    	int checksumsSheetVcmCellNum =0; 
		    	//int checksumsCellNum = 0 ; 
		    	
		    	String sha1 = ""; 
		    	String path = ""; 
		    	
		    	//creo 11 celle in checksums per ogni riga trovata nel tab Grezzi
		    	for(int i=0; i<11;i++) {
		    		 rowChecksums.createCell(i);  
		    	}
		    	
		    	while(cellIterator.hasNext()) {
		    		
		    		Cell cell = cellIterator.next() ;
		    		
		    		switch(cell.getCellType()) {
		    		
			    		case NUMERIC:
			    			break ; 
			    		case STRING:
			    			
			    			if(cellNum == 0) { //faccio la formula per gli sha1
			    				
			    				//sha1 = cell.getStringCellValue() ; 
			    				Cell cellChecksumsUpdate = checksumsSheet.getRow(checksumsRowNum).getCell(2);
			    				String formula = "Grezzi!A"+indexRow;
			    				//cellChecksumsUpdate.setCellValue(sha1); // non devo metterlo direttamente ma fare la formula di excel che me lo prende automaticamente
			    				cellChecksumsUpdate.setCellFormula(formula);
			    				
			    			}else if(cellNum == 1) { //faccio la formula per i path
			    				
			    			   // path = cell.getStringCellValue() ;
			    				Cell cellChecksumsUpdate = checksumsSheet.getRow(checksumsRowNum).getCell(1);
			    				String formula = "Grezzi!B"+indexRow;
			    				//cellChecksumsUpdate.setCellValue(sha1); // non devo metterlo direttamente ma fare la formula di excel che me lo prende automaticamente
			    				cellChecksumsUpdate.setCellFormula(formula);
			    			}
			    			break ; 
		    		}
		    		
		    		cellNum ++ ; 
		    
		    	}
		    	
		    	rowNum ++ ; 
		    	checksumsRowNum ++ ; 
		    	indexRow ++; 
		    }
		    
/*  ---------------------------------------------------------------------------------------------------------------*/		    
		    
		    FormulaEvaluator evaluator = workbook_fp.getCreationHelper().createFormulaEvaluator() ; 
		    Iterator<Row> checkSRowIterator = checksumsSheet.iterator() ; 
		    rowNum = 0; 
		    
			// se è la prima riga la salto perché è il titolo
			if(checkSRowIterator.hasNext()) {
				checkSRowIterator.next() ;
				//System.out.println("DEBUG1...");
				rowNum ++; 
			}
			
			
			//chiamo leggiVCM qui
			//leggiVCM(v) ; 
			
			int conta = 2; 
			int cont = 2; 
			
			while(checkSRowIterator.hasNext()) {
				
				Row row = checkSRowIterator.next(); 
				
				int cellNum = 0;
				
				Iterator<Cell> checkScellIterator = row.cellIterator(); 
				
				while(checkScellIterator.hasNext()) {
					
					Cell cell = checkScellIterator.next();
					
					if(cellNum == 0) { // se la prima colonna faccio la formula per ricavare il .class
						
						String cella = "B"+conta;
						String c = "@";
						String strVuota = "";
						
						//String formula ="RIGHT("+cella+",5)";
						
						//String formula = "RIGHT("+cella+",LEN("+cella+")-FIND("+cella+",SUBSTITUTE("+cella+","+"\"/\" "+","+cella+",LEN("+cella+")-LEN(SUBSTITUTE("+cella+","+"\"/\" "+","+""+"))),1))";
						String formula = "RIGHT("+cella+",LEN("+cella+")-FIND("+cella+",SUBSTITUTE("+cella+","+"\"/\" "+","+cella+",LEN("+cella+")-LEN(SUBSTITUTE("+cella+","+"\"/\" "+","+""+"))),1))";
						cell.setCellFormula(formula);

						conta++;
						
					}else if(cellNum == 3) {
						
						String cella = "B"+cont;
						String matrice = "Description!B2:C139";
						//ora devo caricare i dati della colonna Column1 (tabella sinistra Checksums)
						caricaDescription(descriptionSheet,v); 
						String formula = "VLOOKUP("+cella+","+matrice+",2,FALSE)";
						cell.setCellFormula(formula);
						cont++;
						
					}else if(cellNum == 5) {
						//cell.setCellValue("Maxio");
						
						//questa funzione memorizza dentro degli arraylist i dati di column1 , column2, column3, column4
						//cosi sono facili da recuperarli e incollarli dentro checksums
						
						leggiVCM(v); 
						
						
						
					}
					
					
					cellNum ++; 
				}
				
				rowNum ++; 
			}
		    
		    
			mioFile.close();
			
			//aggiorna il folgio excel
			FileOutputStream out = new FileOutputStream(f);
			workbook_fp.write(out);
			out.close();
			
			indexRow = 0; 
		
			
		}catch(IOException ex) { ex.printStackTrace(); }
		
	}


	
	
}
