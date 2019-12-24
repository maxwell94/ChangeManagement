import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe che gestisce tutto il Change Management. 
 * Ciascun metodo ha un ruolo ben preciso , per esempio:
 * il metodo "appoggioChangedGames()" andrà nel tab "Appoggio Changed Games" 
 * e gestirà tutto il lavoro che dovrà fare in "Appoggio Changed Games" e restituirà
 * il file aggiornato,
 * il metodo "checkSums()" andrà nel tab "Checksums" e gestirà tutto il lavoro che dovrà
 * fare in checksum e restituirà il file aggiornato e così via.
 * Author : Dedou Maximin */

public class ChangeManagement {
	
	/* Il foglio excel che contiene tutto il Change Management */
	private XSSFWorkbook workbook;
	
	/* I vari tabs del foglio excel */
	private XSSFSheet [] tabs;
	
	private String nomeFoglioExcel; 

	/*Costruttore*/
	public ChangeManagement(String nomeFoglioExcel) {
		
		this.workbook = new XSSFWorkbook();
		this.tabs = new XSSFSheet [8];
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
		generaFoglioExcel(nomeFoglioExcel);
		
	}
    
	/* Metodo che restituisce il foglio Excel*/
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}
    
	/* Metodo che modifica il foglio excel*/
	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}
	
	/* Metodo che genera un foglio excel che è il Change Management nella cartella di lavoro 
	 * (in futuro si potrà modificare così da scegliere il percorso in cui si vorrebbe salvarlo) */
	public void generaFoglioExcel(String nome) {
		
		try {
			 
			 if(nome != null) {
				 
				 File f = new File(nome);
				 if(!f.exists()) { // se non esiste lo crea 
					 
					 FileOutputStream out = new FileOutputStream(new File(nome));
					 this.workbook.write(out);
					 out.close();
				 }else {
					 System.out.println("File già esistente!");
				 }

			 }else {
				 System.out.println("Specificare il nome del file!");
			 }
			 
		}catch(Exception ex) { ex.printStackTrace(); }
		
	}
	
	
}
