import kong.unirest.Unirest
import org.apache.commons.lang3.StringUtils
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.commons.io.FileUtils

import java.security.MessageDigest
import java.util.logging.Logger

import java.nio.charset.Charset

@Grab('com.konghq:unirest-java:3.14.5')
@Grab('org.apache.poi:poi-ooxml:5.2.3')
@Grab('org.apache.poi:poi:5.2.3')
@Grab('commons-io:commons-io:2.13.0')
@Grab('org.apache.commons:commons-lang3:3.13.0')

Logger log = Logger.getLogger("")
log.info ("Start ...")

// achtung dl=1 ist wichtig
String source="${args[0]}"
String write="${args[1]}"

String filename = 'files/neu.xlsx'
String mappe = "Aktuell"

URL url = new URI(source).toURL();
URLConnection connection = url.openConnection()

// Download the file
FileUtils.copyURLToFile(connection.getURL(), new File(filename))

String neu = calcDirHash("files")
String alt = FileUtils.readFileToString(new File('old.hash'),Charset.forName('UTF-8'))

if(neu == alt){
    log.info("same")
    System.exit(0)
}

log.info ("not same")

FileUtils.deleteQuietly(new File('old.hash'))
log.info ("old.hash deleted  ...")
FileUtils.writeStringToFile(new File('old.hash'),neu,"utf-8")
log.info ("new old.hash written  ...")

log.info ("update 1 ...")
update(getHeaderFront() + getLines("300",mappe), "39",write)
log.info ("update 2 ...")
update(getHeaderSponsoren() + getLines("400",mappe), "42",write)

FileUtils.deleteQuietly(new File(filename))

static  String getLines(String breite,String mappe) {

    String res = "<h3>Hauptsponsor</h3>";
    for(List l : getLinesX('Aktuell')){
        if(l.size() > 5 && l.get(6).equals("Hauptsponsor")){
            String text = getLine(l.get(0),l.get(10), l.get(9),breite)
            text = StringUtils.removeEnd(text,",")
            res = res + text +"\n"
        }
    }

    res = res + "<h3>Goldsponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).equals("Gold")){
            String text = getLine(l.get(0),l.get(10), l.get(9),breite)
            text = StringUtils.removeEnd(text,",")
            res = res + text +"\n"
        }
    }

    res = res + "<h3>Silbersponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).equals("Silber")){
            String text = getLine(l.get(0),l.get(10), l.get(9),breite)
            text = StringUtils.removeEnd(text,",")
            res = res + text +"\n"
        }
    }

    res = res + "<h3>Bronzesponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).contains("ronze")){
            String text = getLine(l.get(0),l.get(10), l.get(9),breite)
            text = StringUtils.removeEnd(text,",")
            res = res + text +"\n"
        }
    }

    res = res + "<h3>Siegershirt Sponsoring</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).contains("Siegershirt")){
            String text = getLine(l.get(0),l.get(10), l.get(9),breite)
            text = StringUtils.removeEnd(text,",")
            res = res + text +"\n"
        }
    }


    res = res + "<h3>Gönner</h3>"
    for(List l : getLinesX('Donator')){
        if(l.size() > 5 && l.get(6).contains("Gönner")){
            String text = l.get(0)+ ', ' + l.get(1)
            text = StringUtils.removeEnd(text,", ")
            res = res + '<li><span style="font-size: 14px;">'+ text + '&nbsp;</span></span></li>'
        }
    }

    res = res + "<h3>Donatoren</h3>"
    for(List l : getLinesX('Donator')){
        if(l.size() > 5 && l.get(6).contains('Donator')){
            String text = l.get(0)+ ', ' + l.get(1)
            text = StringUtils.removeEnd(text,", ")
            res = res + '<li><span style="font-size: 14px;">'+text + '&nbsp;</span></span></li>'
        }
    }

    res = res + "<h3>Tombolasponsoren</h3>"
    for(List l : getLinesX('Tombola')){

        if("".equals(l.get(0))){
            continue
        }
        String text = l.get(0)+ ', ' + l.get(1)
        text = StringUtils.removeEnd(text,", ")
        res = res + '<li><span style="font-size: 14px;">'+text + '&nbsp;</span></span></li>'

    }
    return res
}

static List getLinesX(String mappe){
    List ret = new ArrayList();
    Workbook workbook
    workbook = new XSSFWorkbook("files/neu.xlsx");
    Iterator<Sheet> sheetIterator = workbook.sheetIterator();
    while (sheetIterator.hasNext()) {
        Sheet sheet = sheetIterator.next();
        if(sheet.getSheetName().toString().contains(mappe)){
            Iterator<Row> rowIterator = sheet.rowIterator();
            rowIterator.next()
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                List l = new ArrayList();
             while (cellIterator.hasNext()){
                 l.add(cellIterator.next().toString())
             }
                ret.add(l)
            }
        }
    }
    return ret
}

static String getHeaderFront() {
    return "<img style=\"margin-top: 5px; margin-bottom: 5px;\" src=\"https://schuelerturnierworb.imgix.net/sponsoren2.png\" /> <br />"
}

static String getHeaderSponsoren() {
    return "<h4>Ein herzliches Dankeschön unseren Sponsoren, Gönnern Donatoren und Inserenten</h4>"
}

static String getLine(String firma, String pic, String link,String breite) {

    if(link.isEmpty() || link.equals("--")){
        return "";
    }

    if(pic.isEmpty() || pic.equals("--")){
        return "";
    }

    pic = pic.replace('pdf','png')

    String lin = link;
    if(!lin.startsWith("https")){
        lin = "https://" + lin
    }
    return "<a alt=\"${firma}\" target=\"_blank\" href=\"${lin}\"><img src=\"https://schuelerturnierworb.imgix.net/${pic}?w=${breite}&ar=4:1&fit=fill&fill=solid&fill-color=white&exp=1&border=3,FFFFFF\" /></a><br />"
}


static void update(String update, String id ,String url) {
    //FileUtils.writeStringToFile(new File(id+"_hallo.html"),update,"utf-8")
    Unirest.post(url).field("id", "${id}").field("text", "${update}").field("submit", "submit").asEmpty()
}


static String calcDirHash(fileDir) {
    def hash = MessageDigest.getInstance("MD5")
    new File(fileDir).eachFileRecurse { file ->
        if (file.isFile()) {
            file.eachByte 4096, { bytes, size ->
                hash.update(bytes, 0 as byte, size);
            }
        }
    }
    return hash.digest().encodeHex() as String
}