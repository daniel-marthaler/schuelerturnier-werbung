import org.apache.poi.xssf.usermodel.*

import javax.mail.*

import javax.mail.internet.*
import javax.activation.*
import java.util.logging.Logger

@Grab('com.konghq:unirest-java:3.14.5')
@Grab('org.apache.poi:poi-ooxml:5.2.3')
@Grab('org.apache.poi:poi:5.2.3')
@Grab('commons-io:commons-io:2.13.0')
@Grab('org.apache.commons:commons-lang3:3.13.0')
@Grab(group='javax.mail', module='mail', version='1.4.7')

Logger log = Logger.getLogger("")
log.info ("Start ...")

def path = "/Users/mad/Desktop/"
def excelFilePath =path + "einladungen.xlsx"

def workbook = new XSSFWorkbook(new FileInputStream(excelFilePath))

def sheet = workbook.getSheetAt(0)

sheet.each { row ->

    String mail = row.getCell(0)
    String anr = row.getCell(1)
    String name = row.getCell(2)

    String anrede = anr + " " + name

    if(!anr.contains("null")){
        Thread.sleep(6000)
        sendEmail("daniel.marthaler@plaintext.ch",anrede, path)
    }

}

workbook.close()

def sendEmail(String empfaenger,String anrede,String path) {

    def password = "***"
    def username = "daniel@marthaler.io"

    def subject = "Einladung zum Start der Plaintext GmbH v1.0 und zum SBB Abschiedsapero von Daniel Marthaler"

    def body = """
<html>
<body>
    <p>"""+anrede+"""</p>

    Du bist herzlich eingeladen:

    <p>Am 11. Januar 2024 um 17:00 beim Freizeithaus Meielen in Zollikofen</p>

    <p><strong>Programm:</strong></p>
    <ul>
        <li>17:00 - 18:30: Abschiedsapero</li>
        <li>18:30 - 23:00: Outdoorfondue / Wurst vom Grill, als Fondue-Alternative</li>
    </ul>

    <p>Der Anlass findet draussen statt, deshalb werden warme Kleider empfohlen. <br/>
    Gen&uuml;gend Platz zum Aufw&auml;rmen ist Freizeithaus ist vorhanden.</p>

    <p>Bitte um An - und Abmeldung bis am 04.12.2023 per E-Mail an: 
    <a href="mailto:daniel@marthaler.io">daniel@marthaler.io</a></p>

    <p>Bei der Anmeldung angeben:</p>
    <ul>
        <li>[  ] Fondue</li>
        <li>[  ] Wurst</li>
    </ul>

    <p>Geschenke: Bitte keine, ein K&auml;sseli wird trotzdem bereitstehen, der Inhalt 
    kommt vollumf&auml;nglich dem <br/>Kinderheim Aeschbacherhus in M&uuml;nsinge zugute.</p>

    <p>Ich freue mich auf einen gem&uuml;tlichen Abend...</p>

    <p>Liebe Gr&uuml;sse<br/>
    D&auml;nu</p><br/>
</body>
</html>
"""

    def properties = new Properties()
    properties.setProperty("mail.smtp.auth", "true")
    properties.setProperty("mail.smtp.starttls.enable", "true")
    properties.setProperty("mail.smtp.starttls.required", "true")

    properties.setProperty("mail.smtp.host", "asmtp.mail.hostpoint.ch")
    properties.setProperty("mail.smtp.port", "587")

    // Set TLS protocol and specific cipher suites
    properties.setProperty("mail.smtp.ssl.protocols", "TLSv1.2")
    properties.setProperty("mail.smtp.ssl.ciphersuites", "TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256")


    def session = Session.getInstance(properties, new javax.mail.Authenticator() {
        protected PasswordAuthentication getPasswordAuthentication() {
            return new PasswordAuthentication(username,  password)
        }
    })

    try {
        def message = new MimeMessage(session)
        message.setFrom(new InternetAddress(username))
        message.setRecipients(Message.RecipientType.TO, empfaenger)
        message.setSubject(subject)



        BodyPart messageBodyPart = new MimeBodyPart()
        messageBodyPart.setContent(body, "text/html")


        BodyPart attachmentBodyPart = new MimeBodyPart()

        String filename = path + "Einladung-11-1-24.png"
        DataSource source = new FileDataSource(filename)
        attachmentBodyPart.setDataHandler(new DataHandler(source))
        attachmentBodyPart.setFileName(filename)


        Multipart multipart = new MimeMultipart()
        multipart.addBodyPart(messageBodyPart)
        multipart.addBodyPart(attachmentBodyPart)

        message.setContent(multipart)

        Transport.send(message)

        println("Email weg zu: " + empfaenger)
    } catch (MessagingException e) {
        e.printStackTrace()
        println("Error sending email: " + e.message)
    }
}

