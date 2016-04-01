/*
    AUTEUR :: SUON Samnang
    CONTACT :: samnangsuon.ss@gmail.com
    DATE :: Mars 2016
*/
package controller;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

public class TrouverCourrielDesManagers {
    
    final static String chemin = JOptionPane.showInputDialog(null, "Veuillez saisir le chemin : ", "Question", JOptionPane.QUESTION_MESSAGE);//"M:\\Bombardier Aerospace (1 of 1)-March 2016-Accounts-Review\\Initial Sent Emails\\";    
    public static String nomClient = JOptionPane.showInputDialog(null, "Veuillez saisir le nom du client : ", "Question", JOptionPane.QUESTION_MESSAGE);    
    final static File folder = new File(chemin);
    
    public static void main (String[] args) throws IOException {        
        listFilesForFolder(folder);
        JOptionPane.showMessageDialog(null, "Mon travail est terminé : veuillez regarder dans l'initial sent email.", "Message", JOptionPane.INFORMATION_MESSAGE);
    }
    public static void listFilesForFolder(final File folder) {
        PrintWriter writer = null;
        try {
            writer = new PrintWriter(chemin + "AMS_" + nomClient + ".txt", "UTF-8");
            for (final File fileEntry : folder.listFiles()) {
                if (fileEntry.isDirectory()) {
                    listFilesForFolder(fileEntry);
                } else {
                    if(new Filename(fileEntry.getName(), '/', '.').extension().equals("msg")) {
                        try {
                            MAPIMessage msg = new MAPIMessage(fileEntry.getAbsolutePath());
                            int indiceAccountsReview = new String(fileEntry.getName()).indexOf("Accounts Review");
                            int indiceMSG = new String(fileEntry.getName()).indexOf(".msg");
                            if( indiceAccountsReview > 1 ) {
                                try {
                                    String managerUID = new String(fileEntry.getName()).substring(indiceAccountsReview + 16, indiceMSG);
                                    managerUID = managerUID.trim();
                                    System.out.println("Manager UID : " + managerUID.replace(" ", "."));                                    
                                    System.out.println("Manager email : " + msg.getRecipientEmailAddress());
                                    writer.println(managerUID.replace(" ", ".") + "," + msg.getRecipientEmailAddress());
                                } catch (ChunkNotFoundException ex) {
                                    Logger.getLogger(TrouverCourrielDesManagers.class.getName()).log(Level.SEVERE, null, ex);
                                }
                            }
                        } catch (IOException ex) {
                            Logger.getLogger(TrouverCourrielDesManagers.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(TrouverCourrielDesManagers.class.getName()).log(Level.SEVERE, null, ex);
        } catch (UnsupportedEncodingException ex) {
            Logger.getLogger(TrouverCourrielDesManagers.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            writer.close();
        }
    }
}
class Filename {
    private String fullPath;
    private char pathSeparator, extensionSeparator;

    public Filename(String str, char sep, char ext) {
        fullPath = str;
        pathSeparator = sep;
        extensionSeparator = ext;
    }
    public String extension() {
        int dot = fullPath.lastIndexOf(extensionSeparator);
        return fullPath.substring(dot + 1);
    }
    public String filename() { // gets filename without extension
        int dot = fullPath.lastIndexOf(extensionSeparator);
        int sep = fullPath.lastIndexOf(pathSeparator);
        return fullPath.substring(sep + 1, dot);
    }
    public String path() {
        int sep = fullPath.lastIndexOf(pathSeparator);
        return fullPath.substring(0, sep);
    }
}
/*
Source de :
    http://www.rgagnon.com/javadetails/java-0613.html
    http://www.java2s.com/Code/Java/File-Input-Output/Getextensionpathandfilename.htm
    http://www.tutorialspoint.com/java/java_string_indexof.htm
    http://stackoverflow.com/questions/2885173/how-to-create-a-file-and-write-to-a-file-in-java
*/