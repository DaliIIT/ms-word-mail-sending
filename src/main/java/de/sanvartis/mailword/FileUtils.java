package de.sanvartis.mailword;

import java.io.File;

import static de.sanvartis.mailword._CONST.PROJECT_NAME;
import static de.sanvartis.mailword._CONST.TEMP_PATH;

public class FileUtils {

    public static String getThreadTempFolderPath() {
        String threadName = Thread.currentThread().getName();
        String tempPath = String.format("%s%s_%s\\", TEMP_PATH, PROJECT_NAME, threadName);
        File file = new File(tempPath);
        boolean bool = file.mkdir();
        return tempPath;
    }

    public static String getFileName(String path){
        int lastIndexOfBreak = path.lastIndexOf("\\") != -1 ? path.lastIndexOf("\\") : path.lastIndexOf("/");
        return lastIndexOfBreak != -1 ? path.substring(lastIndexOfBreak + 1) : path;
    }

    public static String changeFileExtension(String fileName,String extension){
        int lastIndexOfDot = fileName.lastIndexOf(".");
        String ext = extension.indexOf(".") == 0 ? extension : "." + extension;
        return lastIndexOfDot != -1 ? fileName.substring(0,lastIndexOfDot) + ext : fileName + ext;
    }
}
