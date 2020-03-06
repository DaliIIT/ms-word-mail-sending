package de.sanvartis.mailword;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.File;
import java.io.IOException;
import java.util.Properties;

import static de.sanvartis.mailword.DocMailUtils.sendWordMail;
import static de.sanvartis.mailword.DocMailUtils.setRecipients;
import static de.sanvartis.mailword.JacobUtils.convertDocToHtml;

@SpringBootApplication
public class MailwordApplication {


    public static void main(String[] args) throws Exception {
        SpringApplication.run(MailwordApplication.class, args);
        String docPath = "C:\\Users\\medal\\IdeaProjects\\mailword\\src\\main\\resources\\aaa.docx";
        String htmlDocName = "test.html";

//        convertDocToHtml(docPath,htmlDocName);

        Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", "192.168.80.250");
        Session session = Session.getDefaultInstance(properties);
        MimeMessage msg = new MimeMessage(session);
        msg.setFrom(new InternetAddress("wajih.elleuch@sanvartis.de"));
        setRecipients(msg, Message.RecipientType.TO, "mohamedali.jallouli@sanvartis.de");
        msg.setSubject("Test");

        sendWordMail(docPath, msg);
    }

//    private static void sendWordMail(String inputDocPath) {
//        ActiveXComponent oWord = new ActiveXComponent("Word.Application");
//        oWord.setProperty("Visible", new Variant(true));
//        Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();
//        Dispatch oDocument = Dispatch.call(oDocuments, "Open", inputDocPath).toDispatch();
//        //open send mail panel
//        Dispatch oActiveWindow = oWord.getProperty("ActiveWindow").toDispatch();
//        Dispatch.put(oActiveWindow,"EnvelopeVisible",true);
//        // Enable attachments
//        Dispatch oOptions = oWord.getProperty("Options").toDispatch();
//        Dispatch.put(oOptions,"SendMailAttach",true);
//        //
//        Dispatch oActiveDocument = oWord.getProperty("ActiveDocument").toDispatch();
//        Dispatch oMailEnvelope = Dispatch.get(oActiveDocument, "MailEnvelope").toDispatch();
//        Dispatch oItem = Dispatch.get(oMailEnvelope, "Item").toDispatch();
//        Dispatch.put(oItem,"Subject","JACOB");
//        Dispatch oRecipients = Dispatch.get(oItem, "Recipients").toDispatch();
//        Dispatch.call(oRecipients,"add","medalijallouli@gmail.com");
//        Dispatch.call(oItem,"Send");
//    }


}
