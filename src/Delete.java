import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Delete {

	public JFrame frame;
	public JTextField textField;
	public JTextField textField_1;
	public JLabel lbl3 = new JLabel(" ");
	
	

	/**
	 * Create the application.
	 */
	public Delete() {
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
		
		Component verticalStrut = Box.createVerticalStrut(20);
		verticalBox.add(verticalStrut);
		
		JButton btnNewButton = new JButton("Delete");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					ty();
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				
			}
		});
		verticalBox.add(btnNewButton);
		
		JPanel panel_2 = new JPanel();
		frame.getContentPane().add(panel_2, BorderLayout.SOUTH);
		
		
		panel_2.add(lbl3);
	}
public void ty() throws Exception
{
	String fileName = "Record.xlsx";
    String cellContent = textField.getText();
    String cell2=textField_1.getText();
    int rownr=0, colnr = 2;

    FileInputStream input = new FileInputStream(fileName);

    XSSFWorkbook wb = new XSSFWorkbook(input);
    XSSFSheet sheet = wb.getSheetAt(0);

    rownr = findRow(sheet, cellContent,cell2);
    if(rownr!=0){
    sheet=output(sheet, rownr-1, colnr);
    FileOutputStream fo=new FileOutputStream("Record.xlsx");
    wb.write(fo);
    lbl3.setText("Deleted Sucessfully");
    }
    else
    {
    	lbl3.setText("Record not found");
    }
    wb.close();

}
public  XSSFSheet output(XSSFSheet sheet, int rownr, int colnr) throws Exception {
   
	int lastRowNum=sheet.getLastRowNum();
    if(rownr>=0&&rownr<lastRowNum){
        sheet.shiftRows(rownr+1,lastRowNum, -1);
    }
    if(rownr==lastRowNum){
        XSSFRow removingRow=sheet.getRow(rownr);
        if(removingRow!=null){
            sheet.removeRow(removingRow);
        }
   
     
        
    }
    return sheet;
  
}

public static int findRow(XSSFSheet sheet, String cellContent,String cell2) {
    for (Row row : sheet) {
    	
        for (Cell cell : row) {
            
                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)&&row.getCell(1).getRichStringCellValue().getString().trim().equals(cell2)) {
                    
                	return row.getRowNum()+1; 
                    
                }
            
        }
    }               
    return 0;
}

}
