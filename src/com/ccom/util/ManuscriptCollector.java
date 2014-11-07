package com.ccom.util;

import java.io.*;
import java.text.*;
import java.util.*;

import javax.mail.*;
import javax.mail.internet.*;

import org.apache.log4j.*;

public class ManuscriptCollector{
 private MimeMessage mimeMessage = null;
 private String saveAttachPath = "";          //附件下载后的存放目录
 private StringBuffer bodytext = new StringBuffer();
 //存放邮件内容的StringBuffer对象
 private String dateformat = "yyyy-MM-dd HH:mm:ss";    //默认的日前显示格式
 static Logger syslogger  =  null;
 static Logger bizlogger  =  null;
 
/**
 * 构造函数,初始化一个MimeMessage对象
 */
 public ManuscriptCollector(){
	 }
 public ManuscriptCollector(MimeMessage mimeMessage){
	 this.mimeMessage = mimeMessage;
 }

 public void setMimeMessage(MimeMessage mimeMessage){
  this.mimeMessage = mimeMessage;
 }
  
/**
 * 获得发件人的地址和姓名
 */
 public String getFrom()throws Exception{
  InternetAddress address[] = (InternetAddress[])mimeMessage.getFrom();
  String from = address[0].getAddress();
  if(from == null) from="";
   String personal = address[0].getPersonal();
   if(personal == null) personal="";
    String fromaddr = personal+"<"+from+">";
    return from;
 }
/**
 * 获得邮件的收件人，抄送，和密送的地址和姓名，根据所传递的参数的不同
 * "to"----收件人 "cc"---抄送人地址 "bcc"---密送人地址
 */

 public String getMailAddress(String type)throws Exception{
  String mailaddr = "";
  String addtype = type.toUpperCase();
  InternetAddress []address = null;
  if(addtype.equals("TO") || addtype.equals("CC") ||addtype.equals("BCC")){
   if(addtype.equals("TO")){
    address = (InternetAddress[])mimeMessage.getRecipients(Message.RecipientType.TO);
   }else if(addtype.equals("CC")){
    address = (InternetAddress[])mimeMessage.getRecipients(Message.RecipientType.CC);
   }else{
    address = (InternetAddress[])mimeMessage.getRecipients(Message.RecipientType.BCC);
   }
   if(address != null){
    for(int i=0;i<address.length;i++){
     String email=address[i].getAddress();
     if(email==null) email="";
     else{
      email=MimeUtility.decodeText(email);
     }
     String personal=address[i].getPersonal();
      if(personal==null) personal="";
      else{
       personal=MimeUtility.decodeText(personal);
      }
      String compositeto=personal+"<"+email+">";
      mailaddr+=","+compositeto;
     }
     mailaddr=mailaddr.substring(1);
    }
   }else{
   throw new Exception("Error emailaddr type!");
   }
   return mailaddr;
  }
    
 /**
  * 获得邮件主题
  */

  public String getSubject()throws MessagingException{
   String subject = "";
   try{
    subject = MimeUtility.decodeText(mimeMessage.getSubject());
    if(subject == null) subject="";
   }catch(Exception exce){
   }
   return subject;
  }

  /**
   * 获得邮件发送日期
 * @throws MessagingException 
   */

  public void DeleteMessage() throws MessagingException
  {
	  mimeMessage.setFlag(Flags.Flag.DELETED, true);
  }
  
 /**
  * 获得邮件发送日期
  */

  public String getSentDate()throws Exception{
    Date sentdate = mimeMessage.getSentDate();
    SimpleDateFormat format = new SimpleDateFormat(dateformat);
    return format.format(sentdate);
  }

  public Date getSentDate2()throws Exception{
	    Date sentdate = mimeMessage.getSentDate();
	    return sentdate;
	  }
  
  
 /**
  * 获得邮件正文内容
  */

  public String getBodyText(){
   return bodytext.toString();
  }
    
 /**
  * 解析邮件，把得到的邮件内容保存到一个StringBuffer对象中，解析邮件
  * 主要是根据MimeType类型的不同执行不同的操作，一步一步的解析
  */

  public void getMailContent(Part part)throws Exception{
    String contenttype = part.getContentType();
    int nameindex = contenttype.indexOf("name");
    boolean conname =false;
    if(nameindex != -1) conname=true;
//     System.out.println("CONTENTTYPE: "+contenttype);
     if(part.isMimeType("text/plain") && !conname){
      bodytext.append((String)part.getContent());
     }else if(part.isMimeType("text/html") && !conname){
      bodytext.append((String)part.getContent());
     }else if(part.isMimeType("multipart/*")){
      Multipart multipart = (Multipart)part.getContent();
      int counts = multipart.getCount();
      for(int i=0;i<counts;i++){
        getMailContent(multipart.getBodyPart(i));
      }
     }else if(part.isMimeType("message/rfc822")){
      getMailContent((Part)part.getContent());
     }else{}
    }

 /**
  * 判断此邮件是否需要回执，如果需要回执返回"true",否则返回"false"
  */
  public boolean getReplySign()throws MessagingException{
    boolean replysign = false;
    String needreply[] = mimeMessage.getHeader("Disposition-Notification-To");
    if(needreply != null){
     replysign = true;
    }
    return replysign;
  }
    
 /**
  * 获得此邮件的Message-ID
  */
  public String getMessageId()throws MessagingException{
   return mimeMessage.getMessageID();
  }
    
 /**
  * 【判断此邮件是否已读，如果未读返回返回false,反之返回true】
  */
  public boolean isNew()throws MessagingException{
   boolean isnew = false;
   Flags flags = ((Message)mimeMessage).getFlags();
   Flags.Flag []flag = flags.getSystemFlags();
//   System.out.println("flags's length: "+flag.length);
   for(int i=0;i<flag.length;i++){
    if(flag[i] == Flags.Flag.SEEN){
     isnew=true;
     System.out.println("seen Message.......");
     break;
   }
  }
  return isnew;
 }
 
/**
 * 判断此邮件是否包含附件
 */
 public boolean isContainAttach(Part part)throws Exception{
  boolean attachflag = false;
  String contentType = part.getContentType();
  if(part.isMimeType("multipart/*")){
   Multipart mp = (Multipart)part.getContent();
   for(int i=0;i<mp.getCount();i++){
    BodyPart mpart = mp.getBodyPart(i);
    String disposition = mpart.getDisposition();
    if((disposition != null) &&((disposition.equals(Part.ATTACHMENT)) ||(disposition.equals(Part.INLINE))))
     attachflag = true;
    else if(mpart.isMimeType("multipart/*")){
     attachflag = isContainAttach((Part)mpart);
    }else{
     String contype = mpart.getContentType();
     if(contype.toLowerCase().indexOf("application") != -1) attachflag=true;
     if(contype.toLowerCase().indexOf("name") != -1) attachflag=true;
    }
   }
  }else if(part.isMimeType("message/rfc822")){
   attachflag = isContainAttach((Part)part.getContent());
  }
  return attachflag;
 }
   
/**
 * 【保存附件】
 */

 public void saveAttachMent(Part part)throws Exception{
  String fileName = "";
  if(part.isMimeType("multipart/*")){
   Multipart mp = (Multipart)part.getContent();
   for(int i=0;i<mp.getCount();i++){
    BodyPart mpart = mp.getBodyPart(i);
    String disposition = mpart.getDisposition();
    if((disposition != null) &&((disposition.equals(Part.ATTACHMENT)) ||(disposition.equals(Part.INLINE)))){
     fileName = mpart.getFileName();
     if(fileName.toLowerCase().indexOf("gb2312") != -1){
       fileName = MimeUtility.decodeText(fileName);
     }
      saveFile(fileName,mpart.getInputStream());
 }else if(mpart.isMimeType("multipart/*")){
 saveAttachMent(mpart);
 }else{
  fileName = mpart.getFileName();
  if((fileName != null) && (fileName.toLowerCase().indexOf("GB2312") != -1)){
    fileName=MimeUtility.decodeText(fileName);
    saveFile(fileName,mpart.getInputStream());
  }
 }
}
}else if(part.isMimeType("message/rfc822")){
  saveAttachMent((Part)part.getContent());
}
}
    
/**
 * 【设置附件存放路径】
 */

 public void setAttachPath(String attachpath){
  this.saveAttachPath = attachpath;
 }
    
/**
 * 【设置日期显示格式】
 */

 public void setDateFormat(String format)throws Exception{
   this.dateformat = format;
 }
    
/**
 * 【获得附件存放路径】
 */

 public String getAttachPath(){
   return saveAttachPath;
 }
    
/**
 * 【真正的保存附件到指定目录里】
 */

 private void saveFile(String fileName,InputStream in)throws Exception{
  String osName = System.getProperty("os.name");
  String storedir = getAttachPath();
  String separator = "";
  if(osName == null) osName="";
  if(osName.toLowerCase().indexOf("win") != -1){
    separator = "\\";
  if(storedir == null || storedir.equals("")) storedir="c:\\tmp";
  }else{
   separator = "/";
   storedir = "/tmp";
  }
  File storefile = new File(storedir+separator+fileName);
//  System.out.println("storefile's path: "+storefile.toString());
 //for(int i=0;storefile.exists();i++){
 //storefile = new File(storedir+separator+fileName+i);
 //}
 BufferedOutputStream bos = null;
 BufferedInputStream  bis = null;
 try{
  bos = new BufferedOutputStream(new FileOutputStream(storefile));
  bis = new BufferedInputStream(in);
  int c;
  while((c=bis.read()) != -1){
    bos.write(c);
    bos.flush();
  }
 }catch(Exception exception){
   exception.printStackTrace();
   throw new Exception("文件保存失败!");
 }finally{
   bos.close();
   bis.close();
 }
}
   
/**
 * PraseMimeMessage类测试
 */

 public static void main(String args[])throws Exception{
	 
		Properties   props   =  new  Properties();
		try {
			props.load(new FileInputStream("cfg/config.properties"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
	       
		  if(props.isEmpty())
		  {
		   return;
		  }
	 
  String mail_receive_host = props.get("mail_receive_host").toString();
  String mail_send_host = props.get("mail_send_host").toString();
  String mail_username = props.get("mail_username").toString();
  String mail_password = props.get("mail_password").toString();
  String last_max_time = props.get("last_max_time").toString();
  String black_email_list = props.get("black_email_list").toString();
  String my_mail_address = props.get("my_mail_address").toString();
 
  PropertyConfigurator.configure( "cfg/log4j.properties" );
  syslogger  =  Logger.getLogger("SysLog");
  bizlogger  =  Logger.getLogger("BizLog");

  syslogger.info("mail_receive_host:"+mail_receive_host);
  syslogger.info("mail_send_host:"+mail_send_host);
  syslogger.info("mail_username:"+mail_username);
  syslogger.info("mail_password:"+mail_password);
  syslogger.info("black_email_list:"+black_email_list);
  syslogger.info("my_mail_address:"+my_mail_address);
  
  Properties props2 = new Properties();
  props2.put("mail.smtp.host", mail_send_host);
  props2.put("mail.smtp.auth", "true");

  Session session = Session.getDefaultInstance(props2, null);
  Store store = session.getStore("pop3");
  store.connect(mail_receive_host, mail_username, mail_password);
  Folder folder = store.getFolder("INBOX");
  folder.open(Folder.READ_WRITE);
  Message message[] = folder.getMessages();
//  System.out.println("Messages's length: "+message.length);
  bizlogger.info("邮箱中共有["+message.length+"]封邮件");
  bizlogger.info("上次轮询最大时间为:"+last_max_time);
  ManuscriptCollector pmm = null;
  SimpleDateFormat df=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
  Date date_last_max_time=df.parse(last_max_time);
  Date date_current_max_time=date_last_max_time;
  
  for(int i=0;i<message.length;i++){
	   pmm = new ManuscriptCollector((MimeMessage)message[i]);
		  bizlogger.info("开始处理第["+(i+1)+"]封邮件,发送时间为["+pmm.getSentDate()+"]...");
		  String m_sender=pmm.getFrom();
		  if(black_email_list.contains(m_sender))
		  {
			  bizlogger.info("第["+(i+1)+"]封邮件发件人为["+m_sender+"],为垃圾邮件,删除,跳过...");
			  pmm.DeleteMessage();
			  continue;
		  }
		  String m_sentdate=pmm.getSentDate();
		  if(pmm.getSentDate2().getTime() <=date_last_max_time.getTime())
		  {
			  bizlogger.info("第["+(i+1)+"]封邮件发送时间["+m_sentdate+"]小于或等于上次轮询时间["+last_max_time+"],跳过！");
			  continue;
		  }
	  if(date_current_max_time.getTime()<pmm.getSentDate2().getTime())
	       date_current_max_time=pmm.getSentDate2();		  
	  bizlogger.info("第["+(i+1)+"]封邮件发送时间["+m_sentdate+"]大于上次轮询时间["+last_max_time+"],开始处理");
	  String m_subject=pmm.getSubject();
	  bizlogger.info("第["+(i+1)+"]封邮件主题为["+m_subject+"]");
	  if(m_subject.startsWith("[投稿]"))
	  {
		  bizlogger.info("第["+(i+1)+"]封邮件符合规则,开始解析");
		  bizlogger.info("第["+(i+1)+"]封邮件来自["+m_sender+"]");
		  
		  bizlogger.info("开始向["+m_sender+"]发送反馈邮件");
		  MimeMessage sendmessage = new MimeMessage(session);
		  try {
		   // 加载发件人地址
			  sendmessage.setFrom(new InternetAddress(my_mail_address));
		   // 加载收件人地址
			  sendmessage.addRecipients(Message.RecipientType.TO, m_sender);
		   // 加载标题
			  sendmessage.setSubject("收稿反馈");
		   // 向multipart对象中添加邮件的各个部分内容，包括文本内容和附件
		   Multipart multipart = new MimeMultipart();

		   // 设置邮件的文本内容
		   BodyPart contentPart = new MimeBodyPart();
		   contentPart.setText("稿件已收到");
		   multipart.addBodyPart(contentPart);
		   
		   // 将multipart对象放到message中
		   sendmessage.setContent(multipart);
		   // 保存邮件
		   sendmessage.saveChanges();
		   // 发送邮件
		   Transport transport = session.getTransport("smtp");
		   // 连接服务器的邮箱
		   transport.connect(mail_send_host, mail_username, mail_password);
		   // 把邮件发送出去
		   transport.sendMessage(sendmessage, sendmessage.getAllRecipients());
		   transport.close();
		     } catch (Exception e) {
		   e.printStackTrace();
		  }
	  }
	  else
	  {
		  bizlogger.info("第["+(i+1)+"]封邮件不符合规则,记入需要人工核对邮件列表,跳过...");
		  continue;
	  }

	//   System.out.println("Message "+i+" replysign: "+pmm.getReplySign());
	//   System.out.println("Message "+i+" hasRead: "+pmm.isNew());
	//   System.out.println("Message "+i+"  containAttachment: "+pmm.isContainAttach((Part)message[i]));
	//   System.out.println("Message "+i+" form: "+pmm.getFrom());
	//   System.out.println("Message "+i+" to: "+pmm.getMailAddress("to"));
	//   System.out.println("Message "+i+" cc: "+pmm.getMailAddress("cc"));
	//   System.out.println("Message "+i+" bcc: "+pmm.getMailAddress("bcc"));
	//   pmm.setDateFormat("yy年MM月dd日 HH:mm");
	//   System.out.println("Message "+i+" sentdate: "+pmm.getSentDate());
	//   System.out.println("Message "+i+" Message-ID: "+pmm.getMessageId());
	//   pmm.getMailContent((Part)message[i]);
	//   System.out.println("Message "+i+" bodycontent: \r\n"+pmm.getBodyText());
	/* 
	   if(pmm.isContainAttach((Part)message[i])){
	   File dateDir= new File("2014-10-30");
	   if(!dateDir.exists())
	   {
			   dateDir.mkdirs();
	   }   
	   File AttachDir= new File("2014-10-30\\"+pmm.getFrom());
	   if(!AttachDir.exists())
	   {
		   AttachDir.mkdirs();
	   }
	   pmm.setAttachPath(AttachDir.toString());
	   pmm.saveAttachMent((Part)message[i]);
	   }*/
		  bizlogger.info("第["+(i+1)+"]封邮件处理完毕...");

	  }
  
  
  props.setProperty("last_max_time", df.format(date_current_max_time));  
//  props.setProperty("last_max_time", "2012-01-01 00:00:00");  
  props.save(new FileOutputStream("cfg/config.properties"), null);
  folder.close(true);
  System.out.println("done!");
 }
}