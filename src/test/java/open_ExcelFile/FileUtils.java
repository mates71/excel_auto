package open_ExcelFile;

import java.io.File;

public class FileUtils {

	
	public static File getLatestFileFromDir(String dirpath){
		File dir=new File (dirpath);
		File[] files=dir.listFiles();
		if(files==null || files.length==0){
			return null;
			
		}
		
		
		File lastModifiedFile = files[0];
		for(int i =1; i<files.length;i++){
			if(lastModifiedFile.lastModified()< files[i].lastModified()){
				lastModifiedFile=files[i];
			}
		}
		return lastModifiedFile;
	}
}
