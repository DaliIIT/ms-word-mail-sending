package de.sanvartis.mailword;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import static de.sanvartis.mailword.FileUtils.getThreadTempFolderPath;

public class JacobUtils {

    static String convertDocToHtml( String docPath,String htmlDocName){
        String tempFolderPath = getThreadTempFolderPath();
        String htmlPath = tempFolderPath + htmlDocName;
        final int HTML_DOC_TYPE = 10;

        ActiveXComponent oWord = null;
        Dispatch oDocument = null;
        Dispatch oWordBasic = null;

        // ---------------------------------------------------------------

        oWord = new ActiveXComponent( "Word.Application" );
        oWord.setProperty( "Visible", new Variant( false ) );
        oDocument = Dispatch.call( oWord.getProperty( "Documents" ).toDispatch(), "Open", docPath ).toDispatch();
        oWordBasic = Dispatch.call( oWord, "WordBasic" ).getDispatch();
        Dispatch.call( oWordBasic, "FileSaveAs", htmlPath, new Variant( HTML_DOC_TYPE ) );
        Dispatch.call( oDocument, "Close", new Variant( true ) );
        return htmlPath;
    }

}
