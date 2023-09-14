import kong.unirest.Unirest
import org.apache.commons.io.FileUtils
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.commons.io.FileUtils

import java.nio.charset.Charset

@Grab('com.konghq:unirest-java:2.2.00')
@Grab('org.apache.poi:poi-ooxml:5.2.2')
@Grab('org.apache.poi:poi:5.2.2')
@Grab('org.apache.logging.log4j:log4j-to-slf4j:2.8.2')
@Grab('commons-io:commons-io:2.13.0')
@Grab('org.apache.ivy:ivy:2.4.0')


// achtung dl=1 ist wichtig
String url='https://www.dropbox.com/scl/fi/r4jz57v55vts5efqq2q2r/SponsoringWebseite.xlsx?rlkey=s27dyaadbl1zqma6ii7ygwkl5&dl=1'
String filename = 'neu.xlsx'
String mappe = "Aktuell"

while( url ) {
    new URL( url ).openConnection().with { conn ->
        conn.instanceFollowRedirects = false
        url = conn.getHeaderField( "Location" )
        if( !url ) {
            new File( filename ).withOutputStream { out ->
                conn.inputStream.with { inp ->
                    out << inp
                    inp.close()
                }
            }
        }
    }
}

String neu  = FileUtils.readFileToString(new File('neu.xlsx'), Charset.forName('UTF-8'))
String alt = FileUtils.readFileToString(new File('alt.xlsx'),Charset.forName('UTF-8'))


update(getHeaderFront() + getLines("300",mappe), "39");
update(getHeaderSponsoren() + getLines("400",mappe), "42");

FileUtils.deleteQuietly(new File('alt.xlsx'))
FileUtils.moveFile(new File('neu.xlsx'),new File('alt.xlsx'))

static  String getLines(String breite,String mappe) {

    String res = "<h3>Hauptsponsor</h3>";
    for(List l : getLinesX('Aktuell')){
        if(l.size() > 5 && l.get(6).equals("Hauptsponsor")){
            res = res + getLine(l.get(0),l.get(10), l.get(9),breite)+"\n"
        }
    }

    res = res + "<h3>Goldsponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).equals("Gold")){
            res = res + getLine(l.get(0),l.get(10), l.get(9),breite)+"\n"
        }
    }

    res = res + "<h3>Silbersponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).equals("Silber")){
            res = res + getLine(l.get(0),l.get(10), l.get(9),breite)+"\n"
        }
    }

    res = res + "<h3>Bronzesponsoren</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).contains("ronze")){
            res = res + getLine(l.get(0),l.get(10), l.get(9),breite)+"\n"
        }
    }

    res = res + "<h3>Siegershirt Sponsoring</h3>"
    for(List l : getLinesX(mappe)){
        if(l.size() > 5 && l.get(6).contains("Siegershirt")){
            res = res + getLine(l.get(0),l.get(10), l.get(9),breite)+"\n"
        }
    }


    res = res + "<h3>Gönner</h3>"
    for(List l : getLinesX('Donator')){
        if(l.size() > 5 && l.get(6).contains("Gönner")){
            res = res + '<li><span style="font-size: 14px;">'+l.get(0)+ ', ' + l.get(1) + '&nbsp;</span></span></li>'
        }
    }

    res = res + "<h3>Donatoren</h3>"
    for(List l : getLinesX('Donator')){
        if(l.size() > 5 && l.get(6).contains('Donator')){
            res = res + '<li><span style="font-size: 14px;">'+l.get(0)+ ', ' + l.get(1) + '&nbsp;</span></span></li>'
        }
    }

    res = res + "<h3>Tombolasponsoren</h3>"
    for(List l : getLinesX('Tombola')){

        if("".equals(l.get(0))){
            continue
        }
        res = res + '<li><span style="font-size: 14px;">'+l.get(0)+ ', ' + l.get(1) + '&nbsp;</span></span></li>'

    }
    return res
}

static List getLinesX(String mappe){
    List ret = new ArrayList();
    Workbook workbook
    workbook = new XSSFWorkbook("neu.xlsx");
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

static String read(String id) {
    return Unirest.get("https://schuelerturnierworb.ch/werbung_read_848485115.php?id=${id}").asString().body
}

static void update(String update, String id) {
    //FileUtils.writeStringToFile(new File(id+"_hallo.html"),update,"utf-8")
    Unirest.post("https://schuelerturnierworb.ch/werbung_write_848485115.php").field("id", "${id}").field("text", "${update}").field("submit", "submit").asEmpty()
}


