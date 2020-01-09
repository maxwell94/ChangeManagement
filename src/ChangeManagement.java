import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
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
			
			//vado nel tab desiderato che in questo caso è Appoggio Changed Games 
			XSSFSheet desiredSheet = workbook.getSheetAt(5);
			System.out.println("\n Aggiornamento tab: "+desiredSheet.getSheetName()+"...\n");
			
			//iteratore righe
			Iterator<Row> RowIterator = desiredSheet.iterator();
			
			
			//se Ã¨ la prima riga la salto perché è il titolo
			if( RowIterator.hasNext() ) RowIterator.next();
	
			//scorro tutte le righe all'interno del tab
			while(RowIterator.hasNext()) {
				
					//prendo ogni riga
					Row row = RowIterator.next() ; 
					
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
									C_Sha1 = cell.getStringCellValue();
									formula = cell.getCellFormula();
									//formula.replaceAll("[1]", nomeFoglioExcel);
									//System.out.println("formula : "+formula);
									cell.setCellType(CellType.BLANK);
								}
								//System.out.print(cell.getBooleanCellValue()+"\t\t");
								break;
					    }
							
						cellNum ++;
				    }
					
				    cellNum = 0; 
				
			} 
			
			System.out.println("\n Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
			//aggiorna il foglio
			FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		
			
		}catch(Exception ex) { ex.printStackTrace(); }
		
		System.out.println("Done ...");

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
	
	/*metodo che legge un file degli sha e lo carica in excel in grezzi */
	public void grezzi(File f) throws IOException {
		
		if(f.exists()) {
			
			BufferedReader br = new BufferedReader(new FileReader(f)) ; 
	        String st; 
	        char [] appoggio = null; 
	        
	        /* apro il mio foglio excel */
			FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
			XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
			
			//vado nel tab desiderato che in questo caso è Grezzi 
			XSSFSheet desiredSheet = workbook.getSheetAt(0);
			System.out.println("\n Aggiornamento tab: "+desiredSheet.getSheetName()+"...\n");
			
			/* cancello tutti i dati che ci sono già nel tab grezzi */
			deleteSheetAllContent(desiredSheet);
            
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
	        
			//aggiorna il foglio
			FileOutputStream fileOutputStream = new FileOutputStream(new File(nomeFoglioExcel));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			System.out.println("\n Fine aggiornamento tab: "+desiredSheet.getSheetName()+"\n");
			
		}else {
			System.out.println("File non trovato! ");
		}
		
	}
}
