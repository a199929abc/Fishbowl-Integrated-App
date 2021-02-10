

import java.io.*;
import java.net.Socket;

/**
 * Created by IntelliJ IDEA.
 * User: dnewsom
 * Date: 10/10/11
 */
public class Connection {

    private Socket serverSocket;
    private DataOutputStream outputStream;
    private DataInputStream inputStream;

    public Connection() {
        System.out.println("You must connect before you do anything.");
    }

    public Connection(String host, int port) {
        this();
        connect(host, port);
    }

    public Socket getServerSocket() {
        return serverSocket;
    }

    public void setServerSocket(Socket serverSocket) {
        this.serverSocket = serverSocket;
    }

    public DataOutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(DataOutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public DataInputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(DataInputStream inputStream) {
        this.inputStream = inputStream;
    }

    public void connect(String host, int port) {
        try {
            setServerSocket(new Socket(host, port));

            setOutputStream(new DataOutputStream(getServerSocket().getOutputStream()));
            setInputStream(new DataInputStream(getServerSocket().getInputStream()));
            System.out.println("Connection established.");
        } catch (IOException e) {
            System.out.println("Unable to connect to the server.");
            System.exit(1);
        }
    }

    public void disconnect() {
        try {
            getOutputStream().close();
            getInputStream().close();
            getServerSocket().close();
        } catch (IOException e) {
            // do nothing, the program is closing
        }
    }

    /**
     * Sends the xml request to the server to be processed.
     * @param xml the request to send to the server
     * @return the response from the server
     */
    public String sendRequest(String xml) {
        byte[] bytes = xml.getBytes();
        int size = bytes.length;
        try {
            getOutputStream().writeInt(size);
            getOutputStream().write(bytes);
            getOutputStream().flush();
            System.out.println("Request sent.");

            return listenMode();
        } catch (IOException e) {
            System.out.println("The connection to the server was lost.");
            return null;
        }
    }

    private String listenMode() throws IOException {
        byte bytes[];
        int size = getInputStream().readInt();
        bytes = new byte[size];
        getInputStream().readFully(bytes);

        BufferedReader br = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(bytes)));

        String line;
        String msg = "";
        while ((line= br.readLine()) != null) {
            msg = msg + line + "\n";
        }
        br.close();
        return msg;
    }
}
