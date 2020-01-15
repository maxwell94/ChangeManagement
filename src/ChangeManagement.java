import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.util.Iterator;

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
		
//		/*creo i vari tabs del file excel per il Change Management*/
//		this.tabs[0] = this.workbook.createSheet("Grezzi"); // crea un tab Grezzi
//		this.tabs[1] = this.workbook.createSheet("Checksums"); // crea un tab Checksums
//		this.tabs[2] = this.workbook.createSheet("Report"); // crea un tab Report
//		this.tabs[3] = this.workbook.createSheet("Game Versions"); // crea un tab Game Versions
//		this.tabs[4] = this.workbook.createSheet("Changed Games"); // crea un tab Changed Games
//		this.tabs[5] = this.workbook.createSheet("Appoggio Changed Games"); // crea un tab Appoggio Changed Games
//		this.tabs[6] = this.workbook.createSheet("Check EVO"); // crea un tab Check Evo
//		this.tabs[7] = this.workbook.createSheet("Description"); // crea un tab Description
//		
//		//creo fisicamente il foglio excel
		//rt = generaFoglioExcel(nomeFoglioExcel);
		
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

	/* Metodo che lavora all'interno del tab Appoggio Changed Games */
	public void appoggioChangedGames() {
		
		try {
			
			//Apro il foglio excel 
			FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
			XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
			
			//vado nel tab desiderato che in questo caso � Appoggio Changed Games 
			XSSFSheet desiredSheet = workbook.getSheetAt(5);
			System.out.print("\nAggiornamento tab: "+desiredSheet.getSheetName()+"... ");
			
			//iteratore righe
			Iterator<Row> RowIterator = desiredSheet.iterator();
			int rowNum = 0; 
			
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			
			//se � la prima riga la salto perch� � il titolo
			if( RowIterator.hasNext() ) {
				RowIterator.next();
				rowNum ++; 
			}
			
			//scorro tutte le righe all'interno del tab
			while(RowIterator.hasNext()) {
				
					//prendo ogni riga
					Row row = RowIterator.next() ; 
					rowNum ++; 
					
					//scorro tutte le celle
					Iterator<Cell> cellIterator = row.cellIterator();
					int cellNum = 0;
					
					String C_GAME = "";
					String C_FILE = ""; 
					String C_Sha1 = "";
					String formula = ""; 
					
					//scorro tutte le colonne della riga corrente 
					while(cellIterator.hasNext()) {
						
						Cell cell = cellIterator.next(); //prendo ogni cella    
							
						switch(cell.getCellType()) {
							
							case NUMERIC:  
								//System.out.print(cell.getNumericCellValue() +"\t\t");
								break;
							case STRING:
								
								if(cellNum == 0) {
									C_GAME = cell.getStringCellValue();
									//cell.setCellValue("");
									cell.setCellType(CellType.BLANK);
								}else if(cellNum == 1) {
									C_FILE = cell.getStringCellValue();
									cell.setCellType(CellType.BLANK);
								}else if(cellNum == 2) {
									C_Sha1 = cell.getStringCellValue();
									cell.setCellType(CellType.BLANK);
									
								}else if(cellNum == 3) {
								   cell.setCellValue(C_GAME);
								}else if(cellNum == 4){
								   cell.setCellValue(C_FILE);
								}else if(cellNum == 5){
									cell.setCellValue(C_Sha1);
								}
								
								//System.out.print(cell.getStringCellValue()+"\t\t");
								//cell.setCellValue("");
								break;
							case FORMULA:
								
								if(cellNum == 0) {
									C_GAME = cell.getStringCellValue();
									cell.setCellType(CellType.BLANK);
								}else if(cellNum == 1) {
									C_FILE = cell.getStringCellValue();
									cell.setCellType(CellType.BLANK);
								}else if(cellNum == 2) { //C_Sha1
							        switch(cell.getCachedFormulaResultType()) {
							        
						            case NUMERIC:
						            	//ignora
						                //System.out.println("Numerico: " + cell.getNumericCellValue());
						                break;
						            case STRING:
						                //System.out.println("Stringa " + cell.getRichStringCellValue() + "\"");
						            	C_Sha1 = cell.getRichStringCellValue().toString();
						            	cell.setCellType(CellType.BLANK);
						                break;
						            }
									//cell.setCellType(CellType.BLANK);
								}
								//System.out.print(cell.getBooleanCellValue()+"\t\t");
								break;
					    }
							
						cellNum ++;
				    }
					
				    cellNum = 0; 
				
			} 
			
			System.out.print("Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
			//aggiorna il foglio
			FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		
			
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
	
	
	
	/* Metodo che aggiunge una formula in una certa colonna in una precisa cella */
//	public void addFormulaColumn(XSSFSheet desiredSheet , int nColumn , String formula) {
//		
//		Iterator<Row> rowIterator = desiredSheet.iterator();
//		
//		/* Se � la prima riga la salto perch� titolo */
//		if( rowIterator.hasNext() ) {
//			rowIterator.next();
//		}
//		
//		if(formula.length() > 0) { //se c'� almeno qualcosa
//			
//			
//		}else {
//			System.out.println("Formula non valida!");
//		}
//		
//        
//	}
	
	
	
	/* Metodo che lavora all'interno del tab Grezzi , praticamente 
	 * legge il file degli sha e lo carica in excel in grezzi*/
	public void grezzi(File f,String pth) throws IOException, NoSuchAlgorithmException {
		
		if(f.exists()) {
			
			BufferedReader br = new BufferedReader(new FileReader(f)) ; 
	        String st; 
	        String sha1CheckSum = ""; 
	        String nomeCheckSum = ""; 
	        char [] appoggio = null; 
	        
	        /* apro il mio foglio excel */
			FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
			XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
			
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
	        			cell.setCellValue(path);
	        		}
	        	}
	        	 
	        	rowNum ++; 
	        }
	        
	        /* add file prepare_checksums.py*/
	        Row r = desiredSheet.createRow(rowNum++);
	        for(int i=0;i<2;i++) {
	        	Cell c = r.createCell(i);
        		if(i == 0) {
        			c.setCellValue(sha1CheckSum);
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
	public void checksums() throws IOException {
		
        /* apro il mio foglio excel */
		FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
		XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
		
	    //vado nel tab desiderato che in questo caso � Checksums 
		XSSFSheet desiredSheet = workbook.getSheetAt(1);
		System.out.print("\nAggiornamento tab: "+desiredSheet.getSheetName()+"...");
		
		/**
		 * cancello il contenuto di tutte le colonne dalla A alla D nel tab checksum
		 * Quando index=0 � la colonna A , index=1 colonna B , index=2 colonna C ecc. 
		 * E aggiungo le rispettive formule 
		 * */

		deleteContentsAddFormulaColumn(desiredSheet, 1);
		System.out.print(" Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
		
		//aggiorna il foglio
		FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
	}
	
}
