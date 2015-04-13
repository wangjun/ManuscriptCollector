package com.ccom.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.*;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;


public class WordExcelProcessor {

/** Excel 文件要存放的位置，假定在E盘Test目录下*/

public static String outputExcel="fileout/来稿登记.xls",
					 outputExcel2="fileout/其他邮件登记.xls";
private HSSFWorkbook workbook,workbook2;
private HSSFSheet sheet,sheet2;
private FileOutputStream fileOut,fileOut2;
private Configuration configuration = null;
public Logger bizlogger;
public Logger syslogger;

public WordExcelProcessor() {
	configuration = new Configuration();
	configuration.setDefaultEncoding("utf-8");
	PropertyConfigurator.configure( "cfg/log4j.properties" );
    syslogger  =  Logger.getLogger("SysLog");
    bizlogger  =  Logger.getLogger("BizLog");

}


/**
 * 注意dataMap里存放的数据Key值要与模板中的参数相对应
 * @param dataMap
 */
 private void getData(Map<String,Object> dataMap,String sent_time,String name,String title,String phone,String workplace,String email,int index_number)
  {
	  dataMap.put("TITLE", title);
	  dataMap.put("NAME", name);
	  dataMap.put("SENTTIME", sent_time);
	  dataMap.put("INDEX", Integer.toString(index_number));
	  dataMap.put("PHONE", phone);
	  dataMap.put("EMAIL", email);
	  dataMap.put("WORKPLACE", workplace);
  }

 private void getReplyData(Map<String,Object> dataMap,String sent_time,String name,String title,String today)
 {
	  dataMap.put("TITLE", title);
	  dataMap.put("NAME", name);
	  dataMap.put("SENTTIME", sent_time);
	  dataMap.put("TODAY", today);
 } 
 
public void ProcessWord(String sent_time,String name,String title,String phone,String workplace,String email,int index_number){

	//要填入模本的数据文件
	Map<String,Object> dataMap=new HashMap<String,Object>();
	getData(dataMap,sent_time,name,title,phone,workplace,email,index_number);
	//设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
	//这里我们的模板是放在com.havenliu.document.template包下面
//	configuration.setClassForTemplateLoading(this.getClass(), "/com/mail/autoreceiver/template");
	Template t=null;
	try {
		//test.ftl为要装载的模板
		t = configuration.getTemplate("template/1stcheck.ftl");
	} catch (IOException e) {
		e.printStackTrace();
	}
	//输出文档路径及名称
	File outFile = new File("fileout/"+sent_time+"/"+name+"/"+title+"-初审意见书.doc");
	Writer out = null;
	try {
		out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));
	} catch (FileNotFoundException e1) {
		e1.printStackTrace();
	}
	 
    try {
		t.process(dataMap, out);
	} catch (TemplateException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
}

public void ProcessReplyWord(String sent_time,String sent_time2,String name,String title,String today){

	//要填入模本的数据文件
	Map<String,Object> dataMap=new HashMap<String,Object>();
	getReplyData(dataMap,sent_time2,name,title,today);
	//设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
	//这里我们的模板是放在com.havenliu.document.template包下面
//	configuration.setClassForTemplateLoading(this.getClass(), "/com/mail/autoreceiver/template");
	Template t=null;
	try {
		//test.ftl为要装载的模板
		t = configuration.getTemplate("template/reply.ftl");
	} catch (IOException e) {
		e.printStackTrace();
	}
	//输出文档路径及名称
	File outFile = new File("fileout/"+sent_time+"/"+name+"/收稿回复.doc");
	Writer out = null;
	try {
		out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));
	} catch (FileNotFoundException e1) {
		e1.printStackTrace();
	}
	 
    try {
		t.process(dataMap, out);
	} catch (TemplateException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
}

public boolean OpenExcel(){

try{

// 创建新的Excel 工作簿
workbook = new HSSFWorkbook(new FileInputStream(outputExcel));

// 在Excel工作簿中建一工作表，其名为缺省值
// 如要新建一名为"model"的工作表，其语句为：
// HSSFSheet sheet = workbook.createSheet("model");

//HSSFSheet sheet = workbook.createSheet("hhh");
sheet = workbook.getSheetAt(0);
}catch(Exception e) {
System.out.println("已运行 xlCreate() : " + e );
}

// 新建一输出文件流
try {
	fileOut = new FileOutputStream(outputExcel);
} catch (FileNotFoundException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
	return false;
}
return true;
}

public boolean OpenExcel2(){

try{

// 创建新的Excel 工作簿
workbook2 = new HSSFWorkbook(new FileInputStream(outputExcel2));

// 在Excel工作簿中建一工作表，其名为缺省值
// 如要新建一名为"model"的工作表，其语句为：
// HSSFSheet sheet = workbook.createSheet("model");

//HSSFSheet sheet = workbook.createSheet("hhh");
sheet2 = workbook2.getSheetAt(0);
}catch(Exception e) {
System.out.println("已运行 xlCreate() : " + e );
}

// 新建一输出文件流
try {
	fileOut2 = new FileOutputStream(outputExcel2);
} catch (FileNotFoundException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
	return false;
}
return true;
}

public void CloseExcel()
{
	// 操作结束，关闭文件
	try {
		workbook.write(fileOut);
		fileOut.flush();
		fileOut.close();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
}

public void CloseExcel2()
{
	// 操作结束，关闭文件
	try {
		workbook2.write(fileOut2);
		fileOut2.flush();
		fileOut2.close();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
}

public int ProcessExcel(String sent_time,String name,String title,String workplace,String phone,String email){
	
	int index_number=sheet.getLastRowNum()+1;

	// 在索引0的位置创建行（最顶端的行）
	HSSFRow row = sheet.createRow(index_number);

	//在索引0的位置创建单元格（左上端）
	HSSFCell cell = row.createCell((short) 0);


	// 定义单元格为字符串类型
	cell.setCellType(HSSFCell.CELL_TYPE_STRING);


	// 在单元格中输入一些内容
	cell.setCellValue(sent_time);

	cell = row.createCell((short) 1);
	cell.setCellValue(name);

	cell = row.createCell((short) 2);
	cell.setCellValue(title);

	cell = row.createCell((short) 3);
	cell.setCellValue(workplace);

	cell = row.createCell((short) 4);
	cell.setCellValue(phone);

	cell = row.createCell((short) 5);
	cell.setCellValue(email);
	
	SimpleDateFormat df=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	String rec_time=df.format(new Date());
	cell = row.createCell((short) 7);
	cell.setCellValue(rec_time);
	
	return index_number+1;

}

public void ProcessExcel2(String sent_time,String email,String subject){
	
	int index_number=sheet2.getLastRowNum()+1;
	// 在索引0的位置创建行（最顶端的行）
	HSSFRow row = sheet2.createRow(index_number);

	//在索引0的位置创建单元格（左上端）
	HSSFCell cell = row.createCell((short) 0);


	// 定义单元格为字符串类型
	cell.setCellType(HSSFCell.CELL_TYPE_STRING);


	// 在单元格中输入一些内容
	cell.setCellValue(sent_time);

	cell = row.createCell((short) 1);
	cell.setCellValue(email);

	cell = row.createCell((short) 2);
	cell.setCellValue(subject);
	
	SimpleDateFormat df=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	String rec_time=df.format(new Date());
	cell = row.createCell((short) 3);
	cell.setCellValue(rec_time);

}

public void BackUpExcel()
{
	SimpleDateFormat dateformat1=new SimpleDateFormat("yyyy-MM-dd");
	String str_dir=dateformat1.format(new Date());
	File dateDir= new File("backup\\"+str_dir);
	if(!dateDir.exists())
		{
		dateDir.mkdirs();
		}
	 try { 
         int bytesum = 0; 
         int byteread = 0; 
         File oldfile = new File("fileout\\来稿登记.xls"); 
         if (oldfile.exists()) { //文件存在时 
             InputStream inStream = new FileInputStream(oldfile); //读入原文件 
             FileOutputStream fs = new FileOutputStream(dateDir.toString()+"\\来稿登记.xls"); 
             byte[] buffer = new byte[1444]; 
             while ( (byteread = inStream.read(buffer)) != -1) { 
                 bytesum += byteread; //字节数 文件大小 
//                 System.out.println(bytesum); 
                 fs.write(buffer, 0, byteread); 
             } 
             inStream.close(); 
         } 
     } 
     catch (Exception e) { 
         System.out.println("备份\"fileout/来稿登记.xls\"存在问题，请检查是否处于打开状态."); 
         e.printStackTrace(); 

     } 

try { 
         int bytesum = 0; 
         int byteread = 0; 
         File oldfile = new File("fileout\\其他邮件登记.xls"); 
         if (oldfile.exists()) { //文件存在时 
             InputStream inStream = new FileInputStream(oldfile); //读入原文件 
             FileOutputStream fs = new FileOutputStream(dateDir.toString()+"\\其他邮件登记.xls"); 
             byte[] buffer = new byte[1444]; 
             while ( (byteread = inStream.read(buffer)) != -1) { 
                 bytesum += byteread; //字节数 文件大小 
//                 System.out.println(bytesum); 
                 fs.write(buffer, 0, byteread); 
             } 
             inStream.close(); 
         } 
     } 
     catch (Exception e) { 
         System.out.println("备份\"fileout/其他邮件登记.xls\"存在问题，请检查是否处于打开状态."); 
         e.printStackTrace(); 

     } 
}

public boolean SearchExcel2(String keyword){
	boolean isFound=false;
	
	int index_number=sheet2.getLastRowNum();
	// 在索引0的位置创建行（最顶端的行）
	
	for(int i=0 ;i<index_number;i++)
	{
		HSSFRow row = sheet2.getRow(i);
		HSSFCell cell = row.getCell((short) 1);
		if(cell.getStringCellValue().compareTo(keyword)==0)
		{
			isFound=true;
			break;
		}
	}
	
	return isFound;
}

public static void main(String args[]) throws InterruptedException{
//	ProcessExcel();
	
	/*
	CreateExcelWord m_word = new CreateExcelWord();
	m_word.ProcessWord();

       Logger logger  =  Logger.getLogger(CreateExcelWord.class );
       logger.debug( " debug " );
       logger.error( " error " );
       while(true){
       logger.info("hhh");
       ProcessExcel();
       Thread.sleep(5000);
       }
      */ 
	
	
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
	  
	  
	  String mail_host=props.get("mail_host").toString();
	  System.out.println("邮件主机:"+mail_host);
	  PropertyConfigurator.configure( "cfg/log4j.properties" );
      Logger logger  =  Logger.getLogger(WordExcelProcessor.class);
      logger.error( " error " );

}
}