
import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class F {

	public JFrame frame;
	public JTextField textField;
	public JTextField textField_1;
	public JTextField textField_2;
	public JLabel lbl_3 = new JLabel("");

	/**
	 * Launch the application.
	 */
	

	/**
	 * Create the application.
	 */
	public F() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		
		JPanel panel = new JPanel();
		frame.getContentPane().add(panel, BorderLayout.NORTH);
		
		JLabel lblNewLabel = new JLabel("D I Khan Sewa Saniti");
		panel.add(lblNewLabel);
		
		JPanel panel_1 = new JPanel();
		frame.getContentPane().add(panel_1, BorderLayout.CENTER);
		
		Box verticalBox = Box.createVerticalBox();
		panel_1.add(verticalBox);
		
		
		JLabel lblNewLabel_1 = new JLabel("Name");
		verticalBox.add(lblNewLabel_1);
		
		textField = new JTextField();
		verticalBox.add(textField);
		textField.setColumns(10);
		
		JLabel lblNewLabel_2 = new JLabel("Address");
		verticalBox.add(lblNewLabel_2);
		
		textField_1 = new JTextField();
		verticalBox.add(textField_1);
		textField_1.setColumns(10);
		
		JLabel lblCost = new JLabel("Rs.");
		verticalBox.add(lblCost);
		
		textField_2 = new JTextField();
		verticalBox.add(textField_2);
		textField_2.setColumns(10);
		
		Component verticalStrut = Box.createVerticalStrut(20);
		verticalBox.add(verticalStrut);
		
		JButton btnNewButton = new JButton("Add");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String s=textField.getText();
				String y=textField_1.getText();
				String p=textField_2.getText();
				
				try {
					p(s,y,p);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			    
		}});
		verticalBox.add(btnNewButton);
		
		JPanel panel_2 = new JPanel();
		frame.getContentPane().add(panel_2, BorderLayout.SOUTH);
		
		
		panel_2.add(lbl_3);
		
	}
	public void p(String d,String l,String q) throws Exception{
		File f=new File("Record.xlsx");
		if(f.exists()){
			FileInputStream fi=new FileInputStream(f);
			XSSFWorkbook b=new XSSFWorkbook(fi);
			XSSFSheet s =b.getSheetAt(0);
			int k=Delete.findRow(s, d, l);
			if(k==0){
			short v=(short)(s.getLastRowNum()+1);
			XSSFRow row= s.createRow(v);
			row.createCell(0).setCellValue(d);
            row.createCell(1).setCellValue(l);
            row.createCell(2).setCellValue(q);
			fi.close();
			FileOutputStream fo=new FileOutputStream(f);
			b.write(fo);
			b.close();
			fo.close();
			lbl_3.setText("Record Added");
			}
			else
			{
			lbl_3.setText("Record Exist");	
			}
			}
		else
		{
			FileOutputStream fo=new FileOutputStream(f);
			XSSFWorkbook b=new XSSFWorkbook();
			XSSFSheet s =b.createSheet("record1");
			XSSFRow rowhead = s.createRow((short)0);
            rowhead.createCell(0).setCellValue("Name");
            rowhead.createCell(1).setCellValue("Address");
            rowhead.createCell(2).setCellValue("Cost");
            
            XSSFRow row = s.createRow((short)1);
            row.createCell(0).setCellValue(d);
            row.createCell(1).setCellValue(l);
            row.createCell(2).setCellValue(q);
            b.write(fo);
            b.close();
			fo.close();
		
		}
		
	}
}
