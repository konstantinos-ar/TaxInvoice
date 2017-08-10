package com.arvanitis.graphics;

import java.awt.BorderLayout;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;
import java.util.Vector;

import javax.swing.*; 
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Main
{
	private static final String dblink = "jdbc:sqlserver://127.0.0.1:1433";
	private static final String dbuser = "test";
	private static final String dbpass = "test";

	private static final String DBDRIVER = "com.microsoft.sqlserver.jdbc.SQLServerDriver";
	
	static JTable table;
	static File file = new File("C:/Users/user/Downloads/InvoiceTax.txt");

	public static void main(String[] args)
	{
		Statement st = null,st2;
		ResultSet rs = null;
		Connection con = null;
		
		table = new JTable();
		Vector headers = new Vector();
		Vector data2 = new Vector();
		
		HSSFRow row;
		FileInputStream fis = null;
		InputStream urlin = null;
		String data = null;
		int tick = 0, c = 0 ,r = 0;
		String date = null, nav = null, shares = null, assets = null;
		
		
		try
		{
			Class.forName(DBDRIVER);
			con = DriverManager.getConnection(dblink, dbuser, dbpass);

			st = con.createStatement();
			rs = st.executeQuery("Select userid from ZAccess.dbo.Clients where userid like 'a%'");
			
			fis = new FileInputStream(new File("C:/Users/user/Downloads/ACIM_HistoricalNav (3).xls"));

			HSSFWorkbook workbook = new HSSFWorkbook(fis);

			HSSFSheet spreadsheet = workbook.getSheetAt(0);
			headers.clear();
			data2.clear();
			Iterator < Row > rowIterator = spreadsheet.iterator();
			while (rowIterator.hasNext()) 
			{
				row = (HSSFRow) rowIterator.next();
				Iterator < Cell > cellIterator = row.cellIterator();
				c = 0;
				r++;
				Vector d = new Vector();
				while ( cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					c++;
					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_NUMERIC:
						//System.out.print( cell.getNumericCellValue() + " \t\t " );
						data = cell.getStringCellValue();
						if (tick > 0)
						{
							if (c == 2)
								nav = data;
							if (c == 3)
								shares = data;
							if (c == 4)
								assets = data;
							//if (c > 1 && c < 5)
								//d.add(data);
						}
						break;
					case Cell.CELL_TYPE_STRING:
						//System.out.print(cell.getStringCellValue() + " \t\t " );
						data = cell.getStringCellValue();
						if (data.startsWith("Performance"))
							tick = 0;

						if (tick > 0)
						{
							if (c == 1)
								date = data;
							if (c == 2)
								nav = data;
							if (c == 3)
								shares = data;
							if (c == 4)
								assets = data;
							if (c > 0 && c < 5 && r > 4)
								d.add(data);
						}

						if (data.equals("Total Net Assets"))
							tick = 1;
						break;
					}
				}
				if (r > 4 && r < spreadsheet.getLastRowNum()-11){
				//d.add("\n");
				data2.add(d);}
				//System.out.println();
				if (tick > 0 && !data.equals("Total Net Assets"))
				{
					//date = sdf.format(sdf2.parse(date));
					/*try
					{
						System.out.println("Insert into MarketsData.dbo.ETFHist(Sym,Date,Nav,Shares,Assets) values ('','"+date+"',"+nav+","+shares+","+assets+")");
					}
					catch (Exception e){}*/
					//System.out.println("Date: " + date + ", Nav: " + nav + ", Shares: " + shares + ", Assets: " + assets);

				}
			}

			workbook.close();
			//fis.close();
	        
	        System.out.println(data2);
	        headers.clear();
	        headers.add("Date");
	        headers.add("Nav");
	        headers.add("Split");
	        headers.add("Volume");
	        System.out.println(headers);
	        DefaultTableModel model2 = new DefaultTableModel(data2,headers);
	        table.setModel(model2);
	        table.setAutoCreateRowSorter(true);
	        JScrollPane scroll = new JScrollPane(table);
			

			JFrame f = new JFrame();//creating instance of JFrame  
			          
			JButton b = new JButton("Print Sum");//creating instance of JButton  
			b.setBounds(0,0,100, 40);//x axis, y axis, width, height  
			          
			f.add(b, BorderLayout.SOUTH);//adding button in JFrame  
			
			f.add(scroll);
			          
			f.setSize(400,500);//400 width and 500 height
			f.setResizable(true);
			//f.setLayout(null);//using no layout managers  
			f.setVisible(true);//making the frame visible 
	        
			
			TableModel model = table.getModel();
	        FileWriter excel = new FileWriter(file);
	        
	        
	        for(int i = 0; i < model.getColumnCount(); i++){
	            excel.write(model.getColumnName(i) + "\t");
	        }

	        excel.write("\n");

	        for(int i=0; i< model.getRowCount(); i++) {
	            for(int j=0; j < model.getColumnCount(); j++) {
	                excel.write(model.getValueAt(i,j).toString()+"\t");
	            }
	            excel.write("\n");
	        }

	        excel.close();
	        
		}
		catch (IOException | SQLException | ClassNotFoundException e)
		{
			e.printStackTrace();
		}
         
		
	}
}
