
import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.application.Platform;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

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

public class main {
    public static void main(String args []){
    String path ="C:\\Users\\mtcelec2\\Desktop\\RecoveryTemplate.xlsx";

        try {
            Workbook wb=new XSSFWorkbook(new FileInputStream(path));
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
