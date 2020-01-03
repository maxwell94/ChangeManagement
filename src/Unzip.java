import java.awt.List;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Unzip {

    List fileList;
    private String inputFile;
    private String outputFile;
    private BufferedWriter bw;
    FileSha1 fileSha1;
    private static File returnFile ;
    static ArrayList <String> info ;
    
	public Unzip(String inputFile, String outputFile) throws FileNotFoundException {
		this.inputFile = inputFile;
		this.outputFile = outputFile;
		
		this.returnFile = new File("C:\\Users\\Quinel\\Desktop\\QUINEL\\CM Evolution\\CM\\ChangeManagement\\fileSha1.txt");
        FileOutputStream fileOutputStream = new FileOutputStream(returnFile);
        this.bw  = new BufferedWriter(new OutputStreamWriter(fileOutputStream)); 
        
        fileSha1 = new FileSha1() ; 
        info = new ArrayList<String>() ;
	}
	
	public Unzip() {}

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
	

	public ArrayList<String> getInfo() {
		return info;
	}

	public void setInfo(ArrayList<String> info) {
		this.info = info;
	}

	public File getReturnFile() {
		return returnFile;
	}

	public void setReturnFile(File returnFile) {
		this.returnFile = returnFile;
	}

	public BufferedWriter getBw() {
		return bw;
	}

	public void setBw(BufferedWriter bw) {
		this.bw = bw;
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
	 
	 public String rename(String name,String path) {
		 
		 String realName = "";
		 if(name != null && path != null) {
			 
			 realName = estrarre(name);
	 	     File currentName = new File(path+"\\"+name);
	 	     File newName = new File(path+"\\"+realName);
	 	     
	 	     currentName.renameTo(newName);
	 	     
		 }
		return realName;
	 }
	 
	 
	 public static ArrayList<String> salvaInfo() {
		 
		 //System.out.println("stampo: ");
		 BufferedReader objReader = null;
		  try {
		   String strCurrentLine;

		   objReader = new BufferedReader(new FileReader(returnFile));

		   while ((strCurrentLine = objReader.readLine()) != null) {

		      //System.out.println(strCurrentLine);
			  info.add(strCurrentLine); 
		   }

		  } catch (IOException e) {

		   e.printStackTrace();

		  }
		  return info; 
	 }
	 
	 /*Metodo che scorre tutti i file all'interno di una directory rinominata e calcola 
	  * lo Sha1 dei file che trova */
	 public void listf(String directoryName) throws NoSuchAlgorithmException, IOException,FileNotFoundException,NullPointerException {
		    
		    File directory = new File(directoryName);

	        ArrayList<File> resultList = new ArrayList<File>();

	        // get all the files from a directory
	        File[] fList = directory.listFiles();
	        resultList.addAll(Arrays.asList(fList));
	        for (File file : fList) {
	            if (file.isFile()) {
	            	//System.out.println("scrivo: "+fileSha1.sha1Code(file.getAbsolutePath())+" "+file.getAbsolutePath());
	            	bw.write(fileSha1.sha1Code(file.getAbsolutePath())+" "+file.getAbsolutePath()+"\n");
	            } else if (file.isDirectory()) {
	               listf(file.getAbsolutePath());
	            }
	        }
	        
	        /*stampo return file */
	        salvaInfo();
	        
	   }
	 
}
