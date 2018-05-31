/******************************
*Author:Franklin Castellanos
*License:None
*Revision:052618
*******************************/
import java.util.Properties;
import javax.mail.*;
import java.io.IOException;
import javax.mail.internet.MimeBodyPart;

public class EmailAttachment{
	String host;
	String port;
	String protocol;
	String userName;
	String password;
	Properties props;
   int recentFound;  //Index of most recent attachment
   Message[] messages; //All messages from defined folder

   //
	void initServer(String protocol,String host,int port){
      this.props = new Properties();
		props.put("mail.store.protocol", protocol);
		props.put("mail.imap.host", host);
		props.put("mail.imap.port", port);
		props.put("mail.imap.socketFactory" , port );
		props.put("mail.imap.socketFactory.class" , "javax.net.ssl.SSLSocketFactory" );	
	}
   //
   void Login(String userName, String password){
		this.userName = userName;
		this.password = password;
      props.put("mail.imap.user", this.userName);
	}
   
   void StartSession() throws NoSuchProviderException,MessagingException,IOException{
         Session session = Session.getDefaultInstance(props,new javax.mail.Authenticator(){
			   protected PasswordAuthentication getPasswordAuthentication(){
				  return new PasswordAuthentication(userName, password);
			   }
	});
         
        Store store = session.getStore("imaps");
		  store.connect("imap.gmail.com",this.userName,this.password);
		  //Define Folder
		  Folder folder = store.getFolder("inbox");
		  folder.open(Folder.READ_ONLY); 
		  //LOG((String)inbox.getMessageCount());  //get message count
		  messages = folder.getMessages();
      }

   //Default Constructor
	public EmailAttachment(){
	   //Initilalize Email Server
	   initServer("imap","imap.gmail.com",993);
	   //Authenticate
	   Login("f.castellanos@adcommdigitel.com","5levelsof@"); 
	   
          
      try{
        StartSession();
		  int max = messages.length - 1;
		  //LOG(messages[messages.length -1].getSubject());
		  for (int i=max;i>1;i--){
			if(messages[i].getContentType().contains("multipart") && messages[i].getSubject().contains("Schedule") ){
			   Multipart multiPart = (Multipart) messages[i].getContent();
			   int numPart = multiPart.getCount();
			   for (int partCount = 0; partCount < numPart; partCount++) {
				   MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
				   part.saveFile("tst.xls");
				   if((part.getFileName())!=null){
					   LOG(part.getFileName() + " index: " + i);
                this.recentFound = i;
					 System.exit(1);
				  }
			   }
			}
		  }

	  }
		  catch(NoSuchProviderException noProvider){LOG("Provider Exception");}
		  catch(MessagingException mException){mException.printStackTrace();}
		  catch(IOException io){io.printStackTrace();}
	}
// 						if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
// 							this part is attachment
// 							String fileName = part.getFileName();
//                      LOG(messages[i].getSubject());
//                      LOG(String.format("%d",i));
// 							part.saveFile(saveDirectory + File.separator + fileName);
//                   }
	public String getHost(){return this.host;}
	public void setHost(String host){this.host=host;}
	public String getPort(){return this.port;}
	public void setPort(String port){this.port=port;}
   public String getUserName(){return this.userName;}


	void LOG(String arg){
	  System.out.println(arg);
	}
   
   @Override
   public String toString(){
     return(String.format("%s%n",getUserName()) ); 
   }

	 public static void main(String []args){
	   EmailAttachment email = new EmailAttachment();
	 }
}
