import java.util.Properties;
import javax.mail.*;
import java.io.IOException;

public class EmailAttachment{
    String host;
    String port;
    String userName;
    String password;
    
    Properties props = new Properties();
    
    public EmailAttachment(){
       
       //set default server parameters
       props.put("mail.store.protocol", "imap");
       props.put("mail.imap.user", "bloofrank55@gmail.com");
       props.put("mail.imap.host", "imap.gmail.com");
       props.put("mail.imap.port", 993);
       props.put("mail.imap.socketFactory" , 993 );
       props.put("mail.imap.socketFactory.class" , "javax.net.ssl.SSLSocketFactory" );
       
       //Authenticate
       Login("bloofrank55@gmail.com","5Levelsof@"); //use default credentials
       try{
          //Create Session
          Session session = Session.getDefaultInstance(props,new javax.mail.Authenticator(){
               protected PasswordAuthentication getPasswordAuthentication(){
                  return new PasswordAuthentication(userName, password);
               }
              });
          
          Store store = session.getStore("imaps");
          store.connect("imap.gmail.com",this.userName,this.password);  
          
          //Define Folder
          Folder inbox = store.getFolder("inbox");
          inbox.open(Folder.READ_ONLY);
          
          //LOG((String)inbox.getMessageCount());  //get message count
          Message[] messages = inbox.getMessages();
          LOG(messages[messages.length -1].getSubject());
          
          
          }
          catch(NoSuchProviderException noProvider){LOG("Provider Exception");}
          catch(MessagingException mException){mException.printStackTrace();}
          //catch(IOException io){io.printStackTrace();}
    }
    
    protected String getHost(){return this.host;}
    protected void setHost(String host){this.host=host;}
    protected String getPort(){return this.port;}
    protected void setPort(String port){this.port=port;}
    
    protected void Login(String userName, String password){
        this.userName = userName;
        this.password = password;
    }
    
    void LOG(String arg){
      System.out.println(arg);
    }

     public static void main(String []args){
       EmailAttachment email = new EmailAttachment();
        
     }
}