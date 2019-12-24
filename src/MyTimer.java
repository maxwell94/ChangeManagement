import java.sql.Time;
import java.util.Timer;
import java.util.TimerTask;

public class MyTimer extends TimerTask{
    
	DataOra dataOra = new DataOra(); 
	@Override
	public void run() {
		// TODO Auto-generated method stub
		System.out.println(dataOra.DateStamp());
	}
}

