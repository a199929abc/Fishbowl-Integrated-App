package com.fishbowl;

import com.sun.xml.internal.messaging.saaj.packaging.mime.MessagingException;
import com.sun.xml.internal.messaging.saaj.packaging.mime.internet.MimeUtility;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * Created by IntelliJ IDEA.
 * User: dnewsom
 * Date: 10/10/11
 */
public class Requests {

    /**
     * Creates the LoginRq XML.
     * @param userName the username used to login to the server
     * @param userPwd the password used to login to the server
     * @return The XMl login request.
     */
    public static String loginRequest(String userName, String userPwd) {
        String xml = "<FbiXml><Ticket/>" +
                "<FbiMsgsRq>" +
                "<LoginRq>" +
                "<IAID>" + Main.IA_KEY + "</IAID>" +
                "<IAName>" + Main.IA_NAME + "</IAName>" +
                "<IADescription>" + Main.IA_DESC + "</IADescription>" +
                "<UserName>" + userName + "</UserName>" +
                "<UserPassword>" + encodePassword(userPwd) + "</UserPassword>" +
                "</LoginRq>" +
                "</FbiMsgsRq>" +
                "</FbiXml>";

        return xml;
    }

    public static String getSOListRequest(String locationGroup) {
        String xml = "<FbiXml>" +
                "<Ticket>" +
                "<UserID>1</UserID>" +
                "<Key>" + Main.ticketKey + "</Key>" +
                "</Ticket>" +
                "<FbiMsgsRq>" +
                "<GetSOListRq>" +
                "<LocationGroup>" + locationGroup + "</LocationGroup>" +
                "</GetSOListRq>" +
                "</FbiMsgsRq>" +
                "</FbiXml>";

        return xml;
    }

    /**
     * Encrypts the password that will be sent to the server.
     * @param password the password to be encrypted
     * @return The encrypted password
     */
    private static String encodePassword(String password) {
        MessageDigest algorithm;
        try {
            algorithm = MessageDigest.getInstance("MD5");
            algorithm.reset();
            algorithm.update(password.getBytes());
            byte[] encrypted =  algorithm.digest();

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            OutputStream encoder = MimeUtility.encode(out, "base64");

            encoder.write(encrypted);
            encoder.flush();
            return new String(out.toByteArray());
        } catch (NoSuchAlgorithmException e) {
            return "Bad Encryption";
        } catch (MessagingException e) {
            return "Bad Encryption";
        } catch (IOException e) {
            return "Bad Encryption";
        }
    }
}
