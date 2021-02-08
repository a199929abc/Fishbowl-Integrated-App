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
    public static void main(String[] args) {
        String path = "C:\\Users\\mtcelec2\\Desktop\\RecoveryTemplate.xlsx";
 try{
     //创建工作簿
     FileInputStream  is=new FileInputStream(path);
     excelReader reader = new excelReader(is);
     reader.read(is);
     int max_row =reader.maxRow();
    // int title =reader.getMapTitle("Serial #");get mapping title

     for(int i =1;i<max_row;i++){
         ArrayList collect = reader.processRow(i);

         if(collect.get(0)==null&&collect.get(1)==null&&collect.get(2)==null){
             System.out.println("Process have completed！");
             break;
         }
         System.out.println(collect);
     }
     // Create an ArrayList object
     excelWriter ew= new excelWriter(path,"name");
     ew.write_header();
     for(int i=1;i<max_row;i++){
       ArrayList<String> rowFill = new ArrayList<String>();
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         rowFill.add("a");
         ew.write_row(rowFill,i);
     }
     ew.create_excel("C:\\Users\\mtcelec2\\Desktop\\output.xlsx");



 }catch(IOException e){
     e.printStackTrace();
 }
    }

}
