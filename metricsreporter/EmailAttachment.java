package metricsreporter; /******************************
 * *Author:Franklin Castellanos
 * *License:None
 * *Revision:090118
 *******************************/

import javax.mail.*;
import javax.mail.internet.MimeBodyPart;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Properties;

public class EmailAttachment {
    private ArrayList<Message> messages;

    private Properties props = new Properties();
    private String protocol;
    private String host;
    private int port;

    private String userName;
    private String password;

    private long start = 0, stop = 0; //Timer variables

    //Default Constructor
    public EmailAttachment(){
        //Initilalize Email Server
        serverInit("imap", "imap.gmail.com", 993);
        //Authenticate User
        serverLogin("bpsales447@gmail.com", "5levelsof@"); //password for testing only
        try {
            //Begin New Session
            StartSession("inbox");
            //Get Attachment that matches filter
            filterAttachment("Metrics", ".xlsx");
        }
        catch(NoSuchProviderException noProvider) { LOG("Provider Exception"); }
        catch(MessagingException mException){mException.printStackTrace();}
        catch(IOException io){io.printStackTrace();}
    }

    public String getHost(){return this.host;}
    public void setHost(String host) {
        this.host = host;
        props.put("mail.imap.host", host);
    }

    public String getProtocol() {
        return this.protocol;
    }
    public void setProtocol(String protocol) {
        this.protocol = protocol;
        props.put("mail.store.protocol", protocol);
    }

    public int getPort() {
        return this.port;
    }
    public void setPort(int port) {
        this.port = port;
        props.put("mail.imap.port", port);
        props.put("mail.imap.socketFactory", port);
        props.put("mail.imap.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
    }
    public String getUserName(){return this.userName;}

    //Initialize JavaMail server (protocol, host and port)
    void serverInit(String protocol, String host, int port) {
        setProtocol(protocol);
        setHost(host);
        setPort(port);
    }
    //Acquire server serverLogin credentials
    void serverLogin(String userName, String password) {
        this.userName = userName;
        this.password = password;
        //Add username to properties
        props.put("mail.imap.user", this.userName);
    }
    //Start new JavaMail server session within a defined email folder e.g. inbox
    void StartSession(String folder) throws MessagingException {
        Session session = Session.getDefaultInstance(props,new Authenticator(){
            protected PasswordAuthentication getPasswordAuthentication(){
                return new PasswordAuthentication(userName, password);
            }
        });
        Store store = session.getStore(this.getProtocol());
        store.connect("imap.gmail.com", this.getUserName(), this.password);
        //Define Folder
        Folder workingFolder = store.getFolder(folder);
        workingFolder.open(Folder.READ_ONLY);

        messages = new ArrayList<Message>();
        messages.addAll(Arrays.asList(workingFolder.getMessages()));
        //Reverse message order to newest message first
        Collections.reverse(messages); //Adds about 3 secs to search time TODO:optimize
        workingFolder.close(true);
    }

    //Filters attachments based on a defined string and a defined extension.
    //Returns the first message that matches filter
    private MimeBodyPart filterAttachment(String filter, String extension) throws MessagingException, IOException {
        beginTime();
        //Reverse Array before traversal(temporary fix for performance)
        for (Message msg : messages) {
            if (msg.getContent() != null && msg.getSubject() != null) {
                if (msg.getContentType().contains("multipart") && msg.getSubject().contains(filter)) { //If message has attachment
                    Multipart multiPart = (Multipart) msg.getContent();
                    //Traverse Attachment Content
                    for (int partCount = 0; partCount < multiPart.getCount(); partCount++) {
                        MimeBodyPart attachmentPart = (MimeBodyPart) multiPart.getBodyPart(1);
                        if ((attachmentPart.getFileName()) != null && attachmentPart.getFileName().contains(extension)) {
                            attachmentPart.saveFile("recentAttachment.xlsx");
                            LOG(String.format("Found: %s (%s) Saved%n", attachmentPart.getFileName(), msg.getReceivedDate()));
                            endTime();
                            return attachmentPart;
                        }
                    }
                }
            }
        }
        //If no matches are found, return empty Array
        LOG("No Match Found. Try another search.");
        return (new MimeBodyPart());
    }

    //Debug logger
    public void LOG(Object arg){
        System.out.println(arg.toString());
    }
    //Timer
    void beginTime(){
        LOG("Searching...");
        this.start = System.currentTimeMillis();
    }
    void endTime(){
        this.stop = System.currentTimeMillis();
        LOG("Time Spent:" + (stop - start)/1000 + "secs");
    }

    @Override
    public String toString(){ return(String.format("%s%n",getUserName()) ); }

    public static void main(String[] args) {
        EmailAttachment email = new EmailAttachment();
    }

}
