import javax.swing.*;
import java.awt.event.ActionListener;
import java.io.*;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Locale;

import com.sun.scenario.effect.impl.prism.PrImage;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.scene.control.PasswordField;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.stage.FileChooser;
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
import javafx.application.Application;

import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App extends Application {
    public static String userName;
    public static String passWord;
    public static final String IA_NAME = "Java Sample";
    public static final String IA_DESC = "Sample of Java Fishbowl connection";
    public static final int IA_KEY = 54321;
    public static String ticketKey;
    public static String path;
    private static SAXBuilder builder = new SAXBuilder();
    private static String hostName = "universityofvictoria.myfishbowl.com";
    private static int port = 28192;


    public static void main(String[] args) {

        launch();//UI

        Connection connection = new Connection(hostName, port);

        // Create login XML request and send to the server
        String loginRequest = Requests.loginRequest(App.userName,App.passWord);

        String response = connection.sendRequest(loginRequest);
        if (response == null) {
            connection.disconnect();
            System.out.println("response is Null");
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
        /**
         * Request Setup
         *
         *
         *
         *
         * process the excel sheet and write to new file
         * **/

        try {
            //创建工作簿
            FileInputStream is = new FileInputStream(App.path);
            excelReader reader = new excelReader(is);
            reader.read(is);
            int max_row = reader.maxRow();
            int realMaxRow = 0;
            // int title =reader.getMapTitle("Serial #");get mapping title
            excelWriter ew = new excelWriter(path, "name");
            ew.write_header();
            for (int i = 1; i < max_row; i++) {
                ArrayList collect = reader.processRow(i);
                if (collect.get(0) == null && collect.get(1) == null && collect.get(2) == null) {
                    System.out.println("Process have completed！");
                    realMaxRow = i;
                    break;

                }
                if (collect.get(2) != null) {
                    BigInteger partid = new BigDecimal(collect.get(2).toString()).toBigInteger();
                    String getPart = Requests.get_part(partid);
                    response = connection.sendRequest(getPart);
                    String result = getDescription(response);
                    //System.out.println(result.substring( 0,result.indexOf("SN:")-2));//insname
                    //System.out.println(result.substring(result.indexOf("SN:")+4,result.indexOf("DI")));//SN
                    //System.out.println(result.substring(result.indexOf("DI")+4));//DI
                    ArrayList<String> rowFill = new ArrayList<String>();
                    System.out.println("Info are : "+collect);

                    rowFill.add(result.substring(0, result.indexOf("SN:") - 2));
                    rowFill.add(result.substring(result.indexOf("SN:") + 4, result.indexOf("DI")).replace(",", ""));
                    rowFill.add(result.substring(result.indexOf("DI") + 4).replace(",", ""));
                    rowFill.add(partid.toString().replace(",", ""));
                    rowFill.add("Instrument Receiving");
                    rowFill.add("Test and Development");
                    rowFill.add("");
                    rowFill.add("");
                    rowFill.add("");
                    ew.write_row(rowFill, i);
                }//get description
            }
            ew.create_excel("C:\\Users\\mtcelec2\\Desktop\\output.xlsx");


        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    private static Document parseXml(String xml) throws IOException, JDOMException {
        return builder.build(new InputSource(new StringReader(xml)));
    }
    private static String getDescription(String response){
        String result = response.substring(response.indexOf("<Description>")+13,response.indexOf("</Description>"));
        return result;
    }


    @Override
    public void start(Stage primaryStage) throws Exception {
     primaryStage.setTitle("Login");//https://blog.csdn.net/guanguoxiang/article/details/45461879
     GridPane grid =new GridPane();
     grid.setAlignment(Pos.CENTER);
     grid.setHgap(10);
     grid.setVgap(10);
     grid.setPadding(new Insets(25,25,25,25));
     Scene sence =new Scene(grid,350,275);


    Text scenetitle = new Text("Login Fishbowl ");
    scenetitle.setFont(Font.font(20));
    grid.add(scenetitle, 0, 0, 2, 1);



    Label username = new Label("User Name:");
    grid.add(username, 0, 1);
    username.setFont(Font.font(13));


    TextField userTextField = new TextField();
    grid.add(userTextField, 1, 1);

    Label pw = new Label("Password:");
    pw.setFont(Font.font(13));
    grid.add(pw, 0, 2);

    PasswordField pwBox = new PasswordField();
    grid.add(pwBox, 1, 2);

    Button btn_login = new Button("Sign in");
    HBox hbBtn = new HBox(10);

    hbBtn.setAlignment(Pos.BOTTOM_RIGHT);
    hbBtn.getChildren().add(btn_login);
    grid.add(hbBtn, 1, 4);
    Text actiontarget = new Text();
    actiontarget.setFont(Font.font(13));
    grid.add(actiontarget, 1, 6);

    btn_login.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
            actiontarget.setFill(Color.RED);
            actiontarget.setText("Connecting to server . . . . . . ");
            userName=userTextField.getText();
            passWord=pwBox.getText();
            Connection connection = new Connection(hostName, port);
            // Create login XML request and send to the server
            String loginRequest = Requests.loginRequest(userName,passWord);
            String response = connection.sendRequest(loginRequest);
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

            if (statusCode.equals("1000")) {
                //connection build
                connection.disconnect();
                primaryStage.close();
                Stage secondStage = new Stage();
                gotoFile(secondStage);


            }else{
               App.infoBox("username and password doesn't match with system ", "Login Error");
            }

        }
    });
    primaryStage.setScene(sence);
    primaryStage.show();
    }

    private void gotoFile(Stage secondStage) {
        //

        FileChooser fileChooser = new FileChooser();
        fileChooser.setInitialDirectory(new File(System.getProperty("user.home"))
        );
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel", "*.xlsx")

        );
        Text fileChoose = new Text("Please Choose a File ");
        fileChoose.setFont(Font.font(15));

        Button btnProcess =new Button ("Enter");
        Button btnFile = new Button("Explore");
        btnProcess.setMaxSize(150,150);
        btnFile.setMaxSize(150,150);
        GridPane gridPane = new GridPane();
        gridPane.setMinSize(400, 200);
        gridPane.setPadding(new Insets(10, 10, 10, 10));
        //Setting the vertical and horizontal gaps between the columns
        gridPane.setVgap(5);
        gridPane.setHgap(5);
        gridPane.add(fileChoose,0,0);
        gridPane.add(btnFile,0,1);
        gridPane.add(btnProcess,0,2);


        //Setting the Grid alignment
        gridPane.setAlignment(Pos.TOP_LEFT);
        Scene sceneFile = new Scene(gridPane);
        secondStage.setTitle("Open File ");
        btnFile.setOnAction(new EventHandler<ActionEvent>() {
            int i=2;
            public void handle(ActionEvent actionEvent) {
                i+=1;
                FileChooser fileChooser = new FileChooser();
                fileChooser.setInitialDirectory(new File(System.getProperty("user.home"))
                );
                fileChooser.getExtensionFilters().addAll(
                        new FileChooser.ExtensionFilter("Excel", "*.xlsx")

                );
                File file =fileChooser.showOpenDialog(null);
                path = file.getPath();
                Label labelPath1 = new Label(path);
                labelPath1.setTextFill(Color.BLUE);
                System.out.println("path is "+path);
                gridPane.add(labelPath1,0,i);

            }
        });
        btnProcess.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent actionEvent) {

                String fileName = path.substring(path.lastIndexOf("\\")+1);
                infoBox(("Processing file "+fileName),"Successful");
                secondStage.close();




            }
        });


        secondStage.setScene(sceneFile);
        secondStage.show();
    }

    public static void infoBox(String infoMessage, String titleBar)
    {
        JOptionPane.showMessageDialog(null, infoMessage, titleBar, JOptionPane.INFORMATION_MESSAGE);
    }
    }

