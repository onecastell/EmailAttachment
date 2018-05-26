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
	Properties props = new Properties();

	void initServer(String protocol,String host,int port){
		props.put("mail.store.protocol", protocol);
		props.put("mail.imap.host", host);
		props.put("mail.imap.port", port);
		props.put("mail.imap.socketFactory" , port );
		props.put("mail.imap.socketFactory.class" , "javax.net.ssl.SSLSocketFactory" );
		
	}
	
	public EmailAttachment(){
		
	   //set default server parameters
	   initServer("imap","imap.gmail.com",993)
	   props.put("mail.imap.user", "f.castellanos@adcommdigitel.com");
	


	   //Authenticate
	   Login("f.castellanos@adcommdigitel.com","5levelsof@"); //use default credentials
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
		  int max = messages.length - 1;
		  //LOG(messages[messages.length -1].getSubject());
		  for (int i=max;i>1;i--){
			if(messages[i].getContentType().contains("multipart") && messages[i].getSubject().contains("Schedule") ){
			   Multipart multiPart = (Multipart) messages[i].getContent();
			   int numPart = multiPart.getCount();
			   //LOG(Integer.toString(numPart));
			   for (int partCount = 0; partCount < numPart; partCount++) {
						MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
				  //part.saveFile("C:\\Users\\MN2SALES2\\Downloads" + "thisd");
				  if((part.getFileName())!=null){
					 LOG(part.getFileName());
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
