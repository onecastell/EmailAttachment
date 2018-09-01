package metricsreporter; /******************************
 * *Author:Franklin Castellanos
 * *License:None
 * *Revision:090118
 *******************************/

import javax.mail.*;
import javax.mail.internet.MimeBodyPart;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

public class EmailAttachment {
    private String protocol;
    private String host;
    private int port;

    private String userName;
    private String password;

    private Message recentFound;  //Most recently found attachment
    private Message[] messages = {}; //All messages from defined folder
    private Properties props = new Properties();
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

            //Check recently found 1
         /*String recent = getRecent();
         if(recent!=null){
            LOG("Has Recent at: " + recent);
            //System.exit(2);
         }
         else
            LOG("No Recent");*/
            filterAttachment("Metrics", ".xlsx");
        }
        catch(NoSuchProviderException noProvider) { LOG("Provider Exception"); }
        catch(MessagingException mException){mException.printStackTrace();}
        catch(IOException io){io.printStackTrace();}
    }

    public String getHost(){return this.host;}

    //Filters attachments based on a defined string and a defined extension.
    //Returns the first message that matches filter
    private MimeBodyPart filterAttachment(String filter, String extension) throws MessagingException, IOException {
        beginTime();
        for (Message msg : messages) {
            if (msg.getContent() != null && msg.getSubject() != null) {
                if (msg.getContentType().contains("multipart") && msg.getSubject().contains(filter)) { //If message has attachment
                    Multipart multiPart = (Multipart) msg.getContent();
                    //Traverse Attachment Content
                    for (int partCount = 0; partCount < multiPart.getCount(); partCount++) {
                        MimeBodyPart attachmentPart = (MimeBodyPart) multiPart.getBodyPart(1);
                        if ((attachmentPart.getFileName()) != null && attachmentPart.getFileName().contains(extension)) {
                            attachmentPart.saveFile("recentAttachment.xlsx");
                            //setRecent(i);
                            //System.exit(1);//exit program
                            LOG(String.format("Found: %s (%s) Saved%n", attachmentPart.getFileName(), msg.getReceivedDate()));
                            endTime();
                            return attachmentPart;
                        }
                    }
                }
            }
        }
        // if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
        //This part is attachment
        // String fileName = part.getFileName();
        // LOG(messages[i].getSubject());
        // LOG(String.format("%d",i));
        // 	part.saveFile(saveDirectory + File.separator + fileName);
        //}

        //If no matches are found, return empty Array
        LOG("No Match Found. Try another search.");
        return (new MimeBodyPart());
    }

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
    public String getUserName(){return this.userName;}

    public void setPort(int port) {
        this.port = port;
        props.put("mail.imap.port", port);
        props.put("mail.imap.socketFactory", port);
        props.put("mail.imap.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
    }

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
        messages = workingFolder.getMessages();
    }

    //Get index of most recently found attachment
    public String getRecent() {
        String result = null;
        try{
            if (this.recentFound != null) {
                //check local DB File
                BufferedReader reader = new BufferedReader(new FileReader("db.dat"));
                result = reader.readLine();
                reader.close();
            }
            else return null;
        }
        catch(FileNotFoundException fnfE){
            //if no file found
            LOG("DB does not exist");
            return null;
        }
        catch(IOException ioE){ioE.printStackTrace();}
        LOG("Index pulled from file "+ result);
        return result;
    }

    //Debug logger
    public void LOG(Object arg){
        System.out.println(arg.toString());
    }
    //Timer
    void beginTime(){
        LOG("Searching...");
        this.start = System.currentTimeMillis();}
    void endTime(){
        this.stop = System.currentTimeMillis();
        LOG("Time Spent:" + (stop - start)/1000 + "secs");
    }

    @Override
    public String toString(){
        return(String.format("%s%n",getUserName()) );
    }

    //Save index of most recently found attachment
    public void setRecent(Message msg) {
        try {
            this.recentFound = msg;
            //Save to file
            //BufferedWriter writer = new BufferedWriter(new FileWriter("db.dat"));
            //writer.write(String.valueOf(index));//Convert to string for writing
            //writer.close();
        }
        //catch(IOException e){e.printStackTrace();}
        finally {
            LOG("Recent Message Saved");
        }
    }
    public static void main(String[] args) {
        EmailAttachment email = new EmailAttachment();
        //System.out.println(email);
        //xlsFormatter formatter = new xlsFormatter();
        //formatter.openWorkbook("attachment.xls");
        //formatter.getSheetMatrix();
    }

}
