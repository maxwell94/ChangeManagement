import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;

/*Classe che gestisce la data e l'ora corrente */
public class DataOra {
   
	/*attributi*/
	private int giorno; 
	private int mese ; 
	private int anno;
	private Date data;
	private int ora;
	private int minuti;
	String g; 
	String m; 
	String h; 
	String min;
	String aa ; 
	
	/*costruttore*/
	public DataOra() {
		
		data = new Date(); 
		
		this.giorno = data.getDate();
		if(this.giorno < 10) {
			this.g = String.valueOf(this.giorno);
			this.g = "0"+this.g; 
		}else {
			this.g = String.valueOf(this.giorno);
		}
		
		this.mese = data.getMonth()+1;
		if(this.mese < 10) {
			this.m = String.valueOf(this.mese);
			this.m = "0"+this.m; 
		}else {
			this.m = String.valueOf(this.mese);
		}
		
		this.anno = data.getYear();
		
		this.ora = data.getHours();
		
		if(this.ora < 10) {
			this.h = String.valueOf(this.ora);
			this.h = "0"+this.h; 
		}else {
			this.h = String.valueOf(this.ora);
		}
		
		this.minuti = data.getMinutes();
		
		if(this.minuti < 10) {
			this.min = String.valueOf(this.minuti);
			this.min = "0"+this.min; 
		}else {
			this.min = String.valueOf(this.minuti);
		}
		
	} 
	
	public String formatoAnno() {
		String appoggio ="";
		appoggio = Integer.toString(anno);
		appoggio = appoggio.substring(1);
		String aa = "20"+appoggio;
		return aa;
	}
	
	public String DateStamp() {
		
		return "Data: "+g+"/"+m+"/"+formatoAnno()+" "+h+":"+min ; 
	}
	
	
}
