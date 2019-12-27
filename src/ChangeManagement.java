import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe che gestisce tutto il Change Management. 
 * Ciascun metodo ha un ruolo ben preciso , per esempio:
 * il metodo "appoggioChangedGames()" andra nel tab "Appoggio Changed Games" 
 * e gestira tutto il lavoro che dovra fare in "Appoggio Changed Games" e restituir�
 * il file aggiornato,
 * il metodo "checkSums()" andra nel tab "Checksums" e gestira tutto il lavoro che dovr�
 * fare in checksum e restituira il file aggiornato e cosi via.
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
		
		nomeFoglioExcel += ".xlsx";
		this.nomeFoglioExcel = nomeFoglioExcel;
		
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

	/* Metodo che lavora all'interno del tab Appoggio Changed Games */
	public void appoggioChangedGames() {
		
		try {
			
			//Apro il foglio excel 
			FileInputStream fileinputstream = new FileInputStream(new File(nomeFoglioExcel));
			XSSFWorkbook workbook = new XSSFWorkbook(fileinputstream);
			
			//vado nel tab desiderato che in questo caso è Appoggio Changed Games 
			XSSFSheet desiredSheet = workbook.getSheetAt(5);
			System.out.println("\nTab: "+desiredSheet.getSheetName()+"\n");
			
			//iteratore righe
			Iterator<Row> RowIterator = desiredSheet.iterator();
			
			
			//se è la prima riga la salto 
			if( RowIterator.hasNext() ) RowIterator.next();
	
			//scorro tutte le righe all'interno del tab
			while(RowIterator.hasNext()) {
				
					//prendo ogni riga
					Row row = RowIterator.next() ; 
					//scorro tutte le celle
					Iterator<Cell> cellIterator = row.cellIterator();
					int cellNum = 0; 
					
					//scorro tutte le colonne della riga corrente 
					while(cellIterator.hasNext()) {
						
						Cell cell = cellIterator.next(); //prendo ogni cella 
						
						switch(cell.getCellType()) {
							
							case NUMERIC:  
								System.out.print(cell.getNumericCellValue() +"\t\t");
								break;
							case STRING:
								System.out.print(cell.getStringCellValue()+"\t\t");
								break;
							case FORMULA:
								System.out.print(cell.getStringCellValue()+"\t\t");	
					    }
						
						
						cellNum ++;
				    }
					
					System.out.println("");
				    cellNum = 0; 
				
			} 
			
			
		}catch(Exception ex) { ex.printStackTrace(); }

	}
	
}
