import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Locale;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.application.Platform;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;
import org.xml.sax.InputSource;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.xmlbeans.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *  Hello world!
 *
 */
public class App {

    public static final String IA_NAME = "Java Sample";
    public static final String IA_DESC = "Sample of Java Fishbowl connection";
    public static final int IA_KEY = 54321;
    public static String ticketKey;
    private static SAXBuilder builder = new SAXBuilder();

    public static void main(String[] args) {
        Connection connection = new Connection("universityofvictoria.myfishbowl.com", 28192);
        // Create login XML request and send to the server
        String loginRequest = Requests.loginRequest("mtcelec2", "a");

        String response = connection.sendRequest(loginRequest);
        if (response == null) {
            connection.disconnect();
            System.exit(1);
        }

        // Parse response
        Document document = null;
        try {
            document = parseXml(response);
        } catch (Exception e) {
            System.out.println("Cannot parse the xml response.");
            connection.disconnect();
            System.exit(1);
        }
        Element root = document.getRootElement();
        String statusCode = root.getChild("FbiMsgsRs").getAttribute("statusCode").getValue();

        if (!statusCode.equals("1000")) {
            System.out.println("The integrated application needs to be approved.");
            connection.disconnect();
            System.exit(0);
        }
        // Login successful, store the key
        ticketKey = root.getChild("Ticket").getChild("Key").getValue();
        Element node = null;
        String getSOListRequest =Requests.getPart(21689);
        response = connection.sendRequest(getSOListRequest);
        System.out.println(response);
































        String path = "C:\\Users\\kai\\Desktop\\RecoveryTemplate.xlsx";
 try{
     //创建工作簿
     FileInputStream  is=new FileInputStream(path);
     excelReader reader = new excelReader(is);

     reader.read(is);
     int max_row =reader.maxRow();
     int realMaxRow=0;
    // int title =reader.getMapTitle("Serial #");get mapping title

     for(int i =1;i<max_row;i++){
         ArrayList collect = reader.processRow(i);
         if(collect.get(0)==null&&collect.get(1)==null&&collect.get(2)==null){
             System.out.println("Process have completed！");
             realMaxRow=i;
             break;

         }
         System.out.println(collect);
     }
     // Create an ArrayList object
     excelWriter ew= new excelWriter(path,"name");
     ew.write_header();
     for(int i=1;i<realMaxRow;i++){
       ArrayList<String> rowFill = new ArrayList<String>();
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("Instrument Receiving");
         rowFill.add("Test and Development");
         rowFill.add("");
         rowFill.add("");
         rowFill.add("");
         ew.write_row(rowFill,i);
     }
     ew.create_excel("C:\\Users\\kai\\Desktop\\output.xlsx");



 }catch(IOException e){
     e.printStackTrace();
 }
    }  private static Document parseXml(String xml) throws IOException, JDOMException {
        return builder.build(new InputSource(new StringReader(xml)));
    }

}
