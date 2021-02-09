package fishbowlapi;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;
import org.xml.sax.InputSource;
import org.jsoup.nodes.Element;
import java.io.IOException;
import java.io.StringReader;

/**
 * Created by IntelliJ IDEA.
 * User: dnewsom
 * Date: 10/10/11
 */
public class Main {

    public static final String IA_NAME = "Java Sample";
    public static final String IA_DESC = "Sample of Java Fishbowl connection";
    public static final int IA_KEY = 54321;

    public static String ticketKey;

    private static SAXBuilder builder = new SAXBuilder();

    public static void main(String args[]) {
        // Create a connection
        Connection connection = new Connection("localhost", 28192);

        // Create login XML request and send to the server
      
        String loginRequest = com.fishbowl.Requests.loginRequest("mtcelec2", "a");
        
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

        // Verify that the login was successful
        Element root = document.getRootElement();
        String statusCode = root.getChild("FbiMsgsRs").getAttribute("statusCode").getValue();

        if (!statusCode.equals("1000")) {
            System.out.println("The integrated application needs to be approved.");
            connection.disconnect();
            System.exit(0);
        }

        // Login successful, store the key
        ticketKey = root.getChild("Ticket").getChild("Key").getValue();

        // Create the GetSOList XML request and send to the server
        Element node = null;
        String getSOListRequest = com.fishbowl.Requests.getSOListRequest("SLC");
        response = connection.sendRequest(getSOListRequest);

        try {
            document = parseXml(response);
            root = document.getRootElement();
            node = root.getChild("FbiMsgsRs").getChild("GetSOListRs");
        } catch (Exception e) {
            System.out.println("Cannot parse the xml response.");
            connection.disconnect();
            System.exit(1);
        }

        // check error code
        statusCode = node.getAttribute("statusCode").getValue();

        if (statusCode.equals("1000")) {
            node = node.getChild("SalesOrder");
            for (Object obj : node.getChildren()) {
                Element child = (Element) obj;
                if (child.getContentSize() > 1 || !child.getChildren().isEmpty()) {
                    continue;
                }

                System.out.println(child.getName() + " : " + child.getValue());
            }
        } else {
            System.out.println(node.getAttribute("statusMessage").getValue());
        }

        connection.disconnect();
    }

    private static Document parseXml(String xml) throws IOException, JDOMException {
        return builder.build(new InputSource(new StringReader(xml)));
    }
}
