import java.awt.List;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;

public class RenameFolders {

	private String path;
	Unzip unzip;
	private ArrayList<String> allInfo; 

	public RenameFolders(String path) {
		this.path = path;
		unzip = new Unzip();
		allInfo = new ArrayList<String>() ;  
	}
		
    public ArrayList<String> getAllInfo() {
		return allInfo;
	}

	public void setAllInfo(ArrayList<String> allInfo) {
		this.allInfo = allInfo;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}
	
	public void saveAllInfo(ArrayList<String> dati) {
		
		//System.out.println("stampa dati:");
		//System.out.println("n^ elementi : "+dati.size());
		for(int i=0;i<dati.size() ; i++) {
			//System.out.println(dati.get(i));
			allInfo.add(dati.get(i)); 
		}
		//System.out.println("fine stampa");
	}
	
	public void printAllInfo(ArrayList<String> dati) {
		
		System.out.println("stampa dati:");
		System.out.println("n^ elementi : "+dati.size());
		for(int i=0;i<dati.size() ; i++) {
			System.out.println(dati.get(i));
			//allInfo.add(dati.get(i)); 
		}
		System.out.println("fine stampa");
	}
	
	/* Metodo che salva dati di un arrayList su file */
	public void salvaSuFile(ArrayList<String> dati,File f) throws IOException {
		
		if(!dati.isEmpty() && f != null) {
	        FileOutputStream fileOutputStream = new FileOutputStream(f);
	        BufferedWriter bufferedWriter = new BufferedWriter(new OutputStreamWriter(fileOutputStream)); 
	        
			for(int i=0;i<dati.size() ; i++) {
				bufferedWriter.write(dati.get(i)+"\n");
			}
	        
			bufferedWriter.close(); 
		}
	}
	
	public void renameFolders() throws NoSuchAlgorithmException, IOException {
        
        try {
        	
            File mainDirectory = new File (path);
            String name = "";
            String [] folders = null; 
            ArrayList<String> allfiles;
            
            if( mainDirectory.exists() ) {
                 folders = mainDirectory.list(new FilenameFilter() {
                    @Override
                    public boolean accept(File dir, String name) {
                        return new File(dir, name).isDirectory();
                    }
                });
                 
                if( folders.length == 0 ) { //non ha preso niente perché forse sono ancora zippati
                	
                	folders = mainDirectory.list(new FilenameFilter() {
						
						@Override
						public boolean accept(File dir, String name) {
							
						    if (name == null) {
						        return false;
						    }
						    return new File(dir, name).isDirectory() || name.toLowerCase().endsWith(".zip");
						}
					});
                	
                	/* ora scorro gli zip e gli estraggo tutti */
                	for(int i=0;i<folders.length; i++) {
                		//System.out.println(folders[i]);
                		String [] str = folders[i].split(".zip");
                		for(int j=0;j<str.length;j++) {
                			
                			unzip = new Unzip(path+"\\"+folders[i], path+"\\"+str[j]);
                			/*estraggo gli zip nel path corrente*/
							unzip.unZipIt(path+"\\"+folders[i], path+"\\"+str[j]);
							
							/*rinomino le directory*/
							String newName = unzip.rename(str[j],path);
							String newPath = path +"\\"+newName ; //path cartella rinominata
							unzip.listf(newPath);
							
							//chiudo lo stream
							unzip.getBw().close(); 
							allfiles = unzip.salvaInfo() ; 
							saveAllInfo(allfiles);
                		}
                		
                	}
                
                }
                
            }else{
                System.out.println("This directory doesn't exist!");
            }
            
            
        }catch(ArrayIndexOutOfBoundsException ex) {}
	}
    
    
	
}
