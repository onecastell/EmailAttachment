/******************************
*Author:Franklin Castellanos
*License:None
*Revision:052618
*******************************/
import java.util.Properties;
import javax.mail.*;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.FileReader;
import java.io.BufferedWriter;
import java.io.BufferedReader;
import java.io.File;
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
   long start,stop; //Timer variables

   //Default Constructor
	public EmailAttachment(){
	   //Initilalize Email Server
	   initServer("imap","imap.gmail.com",993);
	   //Authenticate
	   Login("f.castellanos@adcommdigitel.com","5levelsof@");           
      try{
         //
         StartSession();
         beginTime();
         //Check recently found 1
         String recent = getRecent();
         if(recent!=null){
            LOG("Has Recent at: " + recent);
            System.exit(2);
         }
         else
            LOG("No Recent");
         
		   int max = messages.length;        
		   for (int i=max;i>1;i--){
			   if(messages[i].getContentType().contains("multipart") && messages[i].getSubject().contains("Schedule") ){
			      Multipart multiPart = (Multipart) messages[i].getContent();
			      int numPart = multiPart.getCount();
			         for (int partCount = 0; partCount < numPart; partCount++) {
				         MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
				         if((part.getFileName())!=null && part.getFileName().contains(".xls")){
                        part.saveFile(part.getFileName());
                        setRecent(i);
                        endTime();
					         System.exit(1);//exit program
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
   
   //Initialize server protocol, host and port
	void initServer(String protocol,String host,int port){
      this.props = new Properties();
		props.put("mail.store.protocol", protocol);
		props.put("mail.imap.host", host);
		props.put("mail.imap.port", port);
		props.put("mail.imap.socketFactory" , port );
		props.put("mail.imap.socketFactory.class" , "javax.net.ssl.SSLSocketFactory" );	
	}
   //Define server login credentials
   void Login(String userName, String password){
		this.userName = userName;
		this.password = password;
      props.put("mail.imap.user", this.userName);
	}
   //Start new email server session
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
		  messages = folder.getMessages();
      }
   //
   public void setRecent(int index){
      try{
         this.recentFound = index;
         //Save to file
         BufferedWriter writer = new BufferedWriter(new FileWriter("db.txt"));
         writer.write(String.valueOf(index));//Convert to string for writing
         writer.close();
      }
      catch(IOException e){e.printStackTrace();}
      finally{LOG(index);}
   }
   //
   public String getRecent() {
      String result = null;
      try{
         if (this.recentFound == 0){
            //check local DB File
            BufferedReader reader = new BufferedReader(new FileReader("db.txt"));
            result = reader.readLine();
            LOG("fetched from file "+ result);
            reader.close();
         }
         else return null;
       }
       catch(FileNotFoundException fnfE){
         //if no file found 
         LOG("DB does not exist");
         return null;
       }
       catch(IOException ioE){}
       finally{return result;}
   }
   
   //Debug logger
	void LOG(Object arg){
	  System.out.println(arg.toString());
	}
   //Timer
   void beginTime(){ this.start = System.currentTimeMillis();}
   void endTime(){
      this.stop = System.currentTimeMillis();
      LOG("Time Spent:" + (stop - start)/1000 + "secs");
   }
   
   @Override
   public String toString(){
     return(String.format("%s%n",getUserName()) ); 
   }

	 public static void main(String []args){
	   EmailAttachment email = new EmailAttachment();
	 }
}
