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
                		//System.out.println(folders[i]);
                		String [] str = folders[i].split(".zip");
                		for(int j=0;j<str.length;j++) {
                			
                			Unzip unzip = new Unzip(path+"\\"+folders[i], path+"\\"+str[j]);
							unzip.unZipIt(path+"\\"+folders[i], path+"\\"+str[j] );
							unzip.rename(str[j],path);
							
                		}
                		
                	}
                }
                
                 
            }else{
                System.out.println("This directory doesn't exist!");
            }
            
            
        }catch(ArrayIndexOutOfBoundsException ex) {}
	}
    
    
	
}
