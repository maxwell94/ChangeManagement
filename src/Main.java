import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.Timer;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		
		String nomeFile = ""; 
		int rt = 0; 
		DataOra dto = new DataOra();
		Scanner scanner = new Scanner(System.in);
				
		System.out.println("---------------------[Change Management]------------------------");
	    System.out.println(dto.DateStamp());
	    System.out.println("Nome foglio excel ? ");
	    nomeFile = scanner.nextLine();
	    if(nomeFile != null) {
	    	//nomeFile += ".xlsx"; 
	        ChangeManagement cm = new ChangeManagement(nomeFile);
	        rt = cm.getRt(); 
	        if( rt == 0) { 
	        	System.out.println("foglio "+nomeFile+" creato con successo");
	        	
	        }else if( rt == -1) { // foglio già esistente quindi pronto per le modifiche
	        	cm.appoggioChangedGames();
	        }
	    	
	    }
	}
	
}
