import java.io.FileInputStream;
import java.io.IOException;
import java.security.DigestInputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

public class FileSha1 {

	public String sha1Code(String filepath) throws IOException,NoSuchAlgorithmException {
		
		FileInputStream fileInputStream = new FileInputStream(filepath) ; 
		MessageDigest messageDigest = MessageDigest.getInstance("SHA-1");
		
		DigestInputStream digestInputStream = new DigestInputStream(fileInputStream, messageDigest);
		byte [] bytes = new byte [1024];
		while(digestInputStream.read(bytes) > 0) ; 
		byte [] resultBytes = messageDigest.digest(); 
		return byteToHexString(resultBytes);
	}

	private String byteToHexString(byte[] bytes) {
		// TODO Auto-generated method stub
		StringBuilder sb = new StringBuilder() ;
		for(byte b: bytes) {
			int value = b & 0xFF;
			
			if(value < 16) {
				sb.append("0");
			}
			
			sb.append(Integer.toHexString(value).toUpperCase());
		}
		return sb.toString();
	}
}
