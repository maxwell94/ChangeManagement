import java.io.File;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;

public class RenameFolders {

	private String path; 

	public RenameFolders(String path) {
		this.path = path; 
	}

    public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public String estrarre(String str){
        int inizio = 16;
        int contaTrattini = 0;
        String strEstratta = "";

        for(int i=16;i<str.length(); i++){

            if(str.charAt(i) == '-'){ // se è uguale '-' controllo se il carattere successivo è un numero
                if( Character.isDigit(str.charAt(i+1)) ){
                    strEstratta = str.substring(inizio,i);
                    break;
                }else{
                    contaTrattini ++;
                }
            }
        }
        return strEstratta;
    }
	
	public void renameFolders() throws FileNotFoundException {
        
        try {
        	
            File mainDirectory = new File (path);
            String name = "";
            String [] folders = null; 
            
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
                		System.out.println(folders[i]);
                		String [] str = folders[i].split(".zip");
                		for(int j=0;j<str.length;j++) {
                			
                			Unzip unzip = new Unzip(folders[i], str[j]);
							unzip.unZipIt(path+"\\"+folders[i], path+"\\"+str[j] );
                			
                		}
                		
                	}
                }
                 
//                for (String nome: folders) {
//                    // chiamo la funzione per estrarre
//                    String realName = estrarre(nome);
//                    System.out.println(realName);
//                    File currentName = new File(path+"/"+nome);
//                    File newName = new File(path+"/"+realName);
//                    
//                    currentName.renameTo(newName);
//                }  
                 
            }else{
                System.out.println("This directory doesn't exist!");
            }
            
            
        }catch(ArrayIndexOutOfBoundsException ex) {}
	}
    
    
	
}
