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
	
	/*costruttore*/
	public DataOra() {
		
		data = new Date(); 
		this.giorno = data.getDate(); 
		this.mese = data.getMonth()+1; 
		this.anno = data.getYear();
		ora = data.getHours();
		minuti = data.getMinutes();
	}

	public int getGiorno() {
		return giorno;
	}

	public void setGiorno(int giorno) {
		this.giorno = giorno;
	}

	public int getMese() {
		return mese;
	}

	public void setMese(int mese) {
		this.mese = mese;
	}

	public int getAnno() {
		return anno;
	}

	public void setAnno(int anno) {
		this.anno = anno;
	} 
	
	public String giornoToString() {
		
		String gg = "";
		
		if(this.giorno == 1) {
			gg = "Lunedi"; 
		}else if(this.giorno == 2) {
			gg = "Martedi";
		}else if(this.giorno == 3) {
			gg = "Mercoledi"; 
		}else if(this.giorno == 4) {
			gg = "Giovedi";
		}else if(this.giorno == 5) {
			gg = "Venerdi";
		}else if(this.giorno == 6) {
			gg = "Sabato";
		}else if(this.giorno == 7) {
			gg = "Domenica";
		}
		return gg; 
	}
	
	public String formatoAnno() {
		String appoggio ="";
		appoggio = Integer.toString(anno);
		appoggio = appoggio.substring(1);
		String aa = "20"+appoggio;
		return aa;
	}
	
	public String DateStamp() {
		
		return "Data: "+giorno+"/"+mese+"/"+formatoAnno()+" "+ora+":"+minuti;
	}
	
	
}
