package de.sanvartis.mailword;

import com.sun.mail.smtp.SMTPTransport;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.IOException;
import java.util.List;

import static de.sanvartis.mailword.FileUtils.*;
import static de.sanvartis.mailword.HtmlUtils.*;
import static de.sanvartis.mailword.JacobUtils.convertDocToHtml;

public class DocMailUtils {

    public static void sendWordMail(String docPath, MimeMessage mail) throws IOException, MessagingException {
        final String HTML_FOLDER_PATH = getThreadTempFolderPath();
        final String HTML_NAME = changeFileExtension(getFileName(docPath),".html");
        String htmlPath = convertDocToHtml(docPath, HTML_NAME);
        String HTML_CONTENT = readFile(htmlPath);
        final List<String> imgs = getImgs(HTML_CONTENT);

        //--------------------------------------------------------------
        final MimeMultipart mimeMultipart = new MimeMultipart();
        for (String imgPath : imgs) {
            final String replacement = "cid:" + getImageName(imgPath);
            HTML_CONTENT = HTML_CONTENT.replaceAll(toUnixPath(imgPath), replacement);

            final MimeBodyPart messageBodyPart = new MimeBodyPart();
            final DataSource imgFile = new FileDataSource(HTML_FOLDER_PATH + imgPath);
            messageBodyPart.setDataHandler(new DataHandler(imgFile));
            messageBodyPart.setHeader("Content-ID", "<" + getImageName(imgPath) + ">");
            mimeMultipart.addBodyPart(messageBodyPart);
        }
        //-----------------------------------------------------------------------

        try {
            final MimeBodyPart mimeBodyPart = new MimeBodyPart();
            mimeBodyPart.setContent(HTML_CONTENT, "text/html");
            mimeMultipart.addBodyPart(mimeBodyPart);
            mail.setContent(mimeMultipart);
//            msg.setFrom(new InternetAddress("andreas.weicher@sanvartis.de"));
            SMTPTransport.send(mail);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void setRecipients(Message email, Message.RecipientType typ, String recipients) throws Exception {

        if (recipients != null) {

            String[] list = new String[1];

            if (!recipients.contains(";"))
                list[0] = recipients;
            else
                list = recipients.split(";");

            for (String address : list)
                if (address.trim().length() > 0)
                    email.addRecipient(typ, new InternetAddress(address.trim()));
        }
    }
}
