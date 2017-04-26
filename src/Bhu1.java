import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
public class Bhu1 {
	static int p,lo,m=1,ty;
	public static void m() throws Exception {
		// TODO Auto-generated method stub
		
		
		File f= new File("g.docx");
		File f2=new File("x.docx");
		File f3=new File("z.docx");
		File f4=new File("Record.xlsx");
		File f5=new File("R.txt");
		FileReader f9 =new FileReader(f5);
		System.out.println((char)0);
		m=f9.read();
		f9.close();
		System.out.println(m);
		String w="File"+m+".docx";
		
		System.out.println(w);
		File n2= new File(w);
		
		FileInputStream fi=new FileInputStream(f4);
		FileWriter f8= new FileWriter(f5);
		XSSFWorkbook b=new XSSFWorkbook(fi);
		XSSFSheet s =b.getSheetAt(0);
		p=s.getPhysicalNumberOfRows();
	     lo=s.getRow(1).getPhysicalNumberOfCells();
		FileOutputStream out = new FileOutputStream(f2);
		 FileInputStream o1 = new FileInputStream(f);
		 XWPFDocument doc= new XWPFDocument(o1);
		 XWPFParagraph paragraph = doc.createParagraph();
	     doc.write(out);
	      doc.close();
	      out.close();
		b.close();
        ty=p;
		f8.write((int)(m+p-1));
		f8.close();
		
		if((p-1)%4==0)
			p=(int)(p/4);
		else
			p=(int)(p/4)+1;
		if(p==1)
			 {
			
		      f2.renameTo(n2);
		      fill(n2,f4);
		
			 }
		else
		{
		      FileOutputStream o = new FileOutputStream(f3);
		      FileInputStream o4 = new FileInputStream(f);
		       FileInputStream o2 = new FileInputStream(f2);
		       merge(o2,o4,o);
		       o4.close();
		       o2.close();
		       o.close();
		       
			f3.renameTo(n2);
			
			fill(n2,f4);
		}
		
		
			}
	public static void merge(InputStream src1, InputStream src2, OutputStream dest) throws Exception {
	    OPCPackage src1Package = OPCPackage.open(src1);
	    OPCPackage src2Package = OPCPackage.open(src2);
	    XWPFDocument src1Document = new XWPFDocument(src1Package);        
	    CTBody src1Body = src1Document.getDocument().getBody();
	    XWPFDocument src2Document = new XWPFDocument(src2Package);
	    CTBody src2Body = src2Document.getDocument().getBody();        
	    appendBody(src1Body, src2Body);
	    src1Document.write(dest);
	}

	private static void appendBody(CTBody src, CTBody append) throws Exception {
	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();
	    String appendString = append.xmlText(optionsOuter);
	    
	    String srcString = src.xmlText();
	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);
	    
	    String mainPart = srcString.substring(srcString.indexOf(">"),srcString.lastIndexOf("<"));
	    String sufix = srcString.substring( srcString.lastIndexOf("<"));
	    String addPart = appendString.substring(appendString.indexOf(">")+1, appendString.lastIndexOf("<"));
	    addPart = concat(addPart,p);
	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
	    src.set(makeBody);
	    
	   }
	private static String concat(String n,int p){
		p=p-1;
		String x=n;
		n="  ";
		while(p!=0){
			n+=x;
			p--;
	}
		return n;
	}
	public static void fill(File fx,File fl) throws Exception
	{
		int l=1;
		FileInputStream i= new FileInputStream(fl);
		XSSFWorkbook t=new XSSFWorkbook(i);
		XSSFSheet s =t.getSheetAt(0);
		while(l<ty)
		{
			XSSFRow r=s.getRow(l);
			
			r.createCell(lo).setCellValue(m);
			String hj=""+m;
			replace(fx,hj,"L");
			m++;
			l++;
		}
		FileOutputStream er= new FileOutputStream(fl);
		t.write(er);
		er.close();
		l=1;
		
		DateFormat dateFormat = new SimpleDateFormat("/mm/yyyy  ");
		Date date = new Date();
		
		String h="15"+dateFormat.format(date);
		while(l<ty)
		{
			XSSFRow r=s.getRow(l);
			String j=r.getCell(0).toString();
			String a=r.getCell(1).toString();
			String q=r.getCell(2).toString();
			String n=r.getCell(3).toString();
			replace(fx,j,"XXX");
			replace(fx,q,"cx");
			replace(fx,a,"XXXX");
			replace(fx,h,"PP");
			replace(fx,h,"PLAN");
			l++;
		}
		
		t.close();
		i.close();
	}
	private static void replace(File fx, String j,String aa) throws Exception{
		FileInputStream f =new FileInputStream(fx);
		XWPFDocument doc = new XWPFDocument(OPCPackage.open(f));
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(aa)) {
		                text = text.replace(aa, j);
		                r.setText(text, 0);
		                f.close();
		        		FileOutputStream f1 =new FileOutputStream(fx);
		        		doc.write(f1);
		        		doc.close();
		        		return;
		        		}
		        }
		    }
		}
		doc.close();
		
}
}
		
		
	
	


