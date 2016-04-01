/*
AUTEUR :: SUON Samnang
CONTACT :: samnangsuon.ss@gmail.com
DATE :: Mars 2016
*/
package controller;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

// You need poi-scratchpad-3.7  and poi-3.7 ( http://poi.apache.org/ )
public class DetectMSGAttachment_1 {
    
    final static String chemin = "M:\\Bombardier Aerospace (1 of 1)-March 2016-Accounts-Review\\Initial Sent Emails\\";
    final static File folder = new File(chemin);//new File("C:\\Users\\samnang.suon\\Desktop\\TEMP\\");
    
    public static void main (String[] args) throws IOException {
        String msgfile = "C:\\Users\\samnang.suon\\Desktop\\TEMP\\AMS   Delete   Disable - Approval for REQ000001064464 - Revue de code d'accès   Account Review - Bombardier Recreatif (1 of 1)  March 2016  erika bially.msg";
        MAPIMessage msg = new MAPIMessage(msgfile);
        AttachmentChunks attachments[] = msg.getAttachmentFiles();
        if(attachments.length > 0) {
            for (AttachmentChunks a  : attachments) {
                System.out.println(a.attachLongFileName);
                // extract attachment
                ByteArrayInputStream fileIn = new ByteArrayInputStream(a.attachData.getValue());
                File f = new File("c:/temp", a.attachLongFileName.toString()); // output
                OutputStream fileOut = null;
                try {
                    fileOut = new FileOutputStream(f);
                    byte[] buffer = new byte[2048];
                    int bNum = fileIn.read(buffer);
                    while(bNum > 0) {
                        fileOut.write(buffer);
                        bNum = fileIn.read(buffer);
                    }
                }
                finally {
                    try {
                        if(fileIn != null) {
                            fileIn.close();
                        }
                    }
                    finally {
                        if(fileOut != null) {
                            fileOut.close();
                        }
                    }
                }
            }
        }
        else {
            System.out.println("No attachment");
        }
        try {
            System.out.println("Recipient Email adresse : " + msg.getRecipientEmailAddress());
            System.out.println("Subject : " + msg.getSubject());
        } catch (ChunkNotFoundException ex) {
            Logger.getLogger(DetectMSGAttachment_1.class.getName()).log(Level.SEVERE, null, ex);
        }        
        
        System.out.println("+++++++++++++++++++++++++++++++++++++++++++++");
        listFilesForFolder(folder);
    }
    public static void listFilesForFolder(final File folder) {
        PrintWriter writer = null;
        try {
            writer = new PrintWriter(chemin + "AMS_Comma Seperated File.txt", "UTF-8");
            for (final File fileEntry : folder.listFiles()) {
                if (fileEntry.isDirectory()) {
                    listFilesForFolder(fileEntry);
                } else {
                    /*
                    System.out.println("Name : " + fileEntry.getName());
                    System.out.println("2016 index : " + new String(fileEntry.getName()).indexOf("2016"));
                    System.out.println("Extension : " + new Filename(fileEntry.getAbsolutePath(), '/', '.').extension());
                    */
                    if(new Filename(fileEntry.getName(), '/', '.').extension().equals("msg")) {
                        try {
                            MAPIMessage msg = new MAPIMessage(fileEntry.getAbsolutePath());
                            int indice = new String(fileEntry.getName()).indexOf("Accounts Review");
                            if( indice > 1 ) {
                                try {
                                    System.out.println("Manager UID : " + new String(fileEntry.getName()).substring(indice + 16, fileEntry.getName().length()));                                    
                                    System.out.println("Manager email : " + msg.getRecipientEmailAddress());
                                    writer.println(new String(fileEntry.getName()).substring(indice + 16, fileEntry.getName().length()) + "," + msg.getRecipientEmailAddress());
                                } catch (ChunkNotFoundException ex) {
                                    Logger.getLogger(DetectMSGAttachment_1.class.getName()).log(Level.SEVERE, null, ex);
                                }
                            }
                        } catch (IOException ex) {
                            Logger.getLogger(DetectMSGAttachment_1.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DetectMSGAttachment_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (UnsupportedEncodingException ex) {
            Logger.getLogger(DetectMSGAttachment_1.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            writer.close();
        }
    }
}
class Filename_1 {
  private String fullPath;
  private char pathSeparator, extensionSeparator;

  public Filename_1(String str, char sep, char ext) {
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