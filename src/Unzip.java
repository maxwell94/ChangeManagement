import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Unzip {

    List fileList;
    private String inputFile;
    private String outputFile;
    
	public Unzip(String inputFile, String outputFile) {
		this.inputFile = inputFile;
		this.outputFile = outputFile;
	}

	public String getInputFile() {
		return inputFile;
	}

	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public String getOutputFile() {
		return outputFile;
	}

	public void setOutputFile(String outputFile) {
		this.outputFile = outputFile;
	}
    
	
	 public void unZipIt(String zipFile, String outputFolder){

	     byte[] buffer = new byte[1024];
	    	
	     try{
	    		
	    	//create output directory is not exists
	    	File folder = new File(outputFile);
	    	if(!folder.exists()){
	    		folder.mkdir();
	    	}
	    		
	    	//get the zip file content
	    	ZipInputStream zis = 
	    		new ZipInputStream(new FileInputStream(zipFile));
	    	//get the zipped file list entry
	    	ZipEntry ze = zis.getNextEntry();
	    		
	    	while(ze!=null){
	    			
	    	   String fileName = ze.getName();
	           File newFile = new File(outputFolder + File.separator + fileName);
	                
	           System.out.println("file unzip : "+ newFile.getAbsoluteFile());
	                
	            //create all non exists folders
	            //else you will hit FileNotFoundException for compressed folder
	            new File(newFile.getParent()).mkdirs();
	              
	            FileOutputStream fos = new FileOutputStream(newFile);             

	            int len;
	            while ((len = zis.read(buffer)) > 0) {
	       		fos.write(buffer, 0, len);
	            }
	        		
	            fos.close();   
	            ze = zis.getNextEntry();
	    	}
	    	
	        zis.closeEntry();
	    	zis.close();
	    		
	    	System.out.println("Done");
	    		
	    }catch(IOException ex){
	       ex.printStackTrace(); 
	    }
	     
	 }
	 
	public String estrarre(String str){

	        int inizio = 16;
	        int contaTrattini = 0;
	        String strEstratta = "";

	        for(int i=16;i<str.length(); i++){

	            if(str.charAt(i) == '-'){ // se è uguale '-' controllo se il carattere successivo è un numero
	                if( Character.isDigit(str.charAt(i+1)) || str.charAt(i+1) == '_' ){
	                    strEstratta = str.substring(inizio,i);
	                    break;
	                }else{
	                    contaTrattini ++;
	                }
	            }
	        }
	        return strEstratta;
	 }
	 
	 public void rename(String name,String path) {
		 
		 if(name != null && path != null) {
			 
			 String realName = estrarre(name);
	 	     File currentName = new File(path+"\\"+name);
	 	     File newName = new File(path+"\\"+realName);
	 	     
	 	     currentName.renameTo(newName);
		 }
	 }
	 
	 /* funzione*/
    
    
}
