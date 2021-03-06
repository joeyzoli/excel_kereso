package excel_kereso;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.FlowLayout;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.ListModel;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListDataListener;
import javax.swing.text.html.HTMLDocument.Iterator;
import javax.swing.JLabel;
import java.awt.event.WindowStateListener;
import java.awt.event.WindowEvent;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class keresoablak 
{

	private JFrame frame;
	private JTextField keresomezo;
	private JList<String> lista;
	private ArrayList<String> eredmeny;
	DefaultListModel<String> model;
	private JTextField tabla_szam;
	private String tablaszam = "4";
	private File f;
	private File[] files;
	private ArrayList<String> szurtfajlok;
	static int i = 0;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) 
	{
		EventQueue.invokeLater(new Runnable() 
		{
			public void run() 
			{
				try 
				{
					keresoablak window = new keresoablak();						//ablak l??trehoz??sa
					window.frame.setVisible(true);								//ablak l??that??v?? t??tele
				} 
				catch (Exception e) 
				{
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public keresoablak() 
	{
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() 
	{
		frame = new JFrame();
		frame.addWindowStateListener(new WindowStateListener() {
			public void windowStateChanged(WindowEvent e) {
			}
		});
		frame.setBounds(100, 100, 640, 480);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setTitle("Keres?? ablak");
		
		JButton keresogomb = new JButton("Keres??s");
		keresogomb.setBounds(445, 63, 89, 23);
		keresogomb.addActionListener(new Kereses());
		frame.getContentPane().add(keresogomb);
		
		model = new DefaultListModel<String>();											//lista modell l??terehoz??sa
		lista = new JList<String>();
		lista.setBounds(127, 167, 407, 186);
		lista.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);					//1x-es kijel??l??s be??ll??t??sa
		lista.setModel(model);															//listamodell be??ll??t??sa
		frame.getContentPane().add(lista);
		
		JScrollPane scrollPane = new JScrollPane(lista);								//scrollozhat?? ablak l??trehoz??sa a Jlistb??l
		scrollPane.setBounds(127, 167, 407, 186);
		frame.getContentPane().add(scrollPane);
		
		keresomezo = new JTextField();
		keresomezo.setBounds(127, 64, 267, 20);
		frame.getContentPane().add(keresomezo);
		keresomezo.setColumns(10);
		
		JLabel keresettszo = new JLabel("Keresend?? sz??:");
		keresettszo.setBounds(127, 19, 131, 34);
		frame.getContentPane().add(keresettszo);
		
		JLabel lblNewLabel = new JLabel("Tal??latok:");
		lblNewLabel.setBounds(127, 142, 81, 14);
		frame.getContentPane().add(lblNewLabel);
		
		JLabel tabla_szama = new JLabel("T??bla sz??ma");
		tabla_szama.setBounds(127, 95, 95, 14);
		frame.getContentPane().add(tabla_szama);
		
		tabla_szam = new JTextField();
		tabla_szam.setText(tablaszam);
		tabla_szam.setBounds(217, 92, 33, 20);
		frame.getContentPane().add(tabla_szam);
		tabla_szam.setColumns(10);
		
		JButton export = new JButton("Export??l??s");
		export.setBounds(445, 389, 89, 23);
		export.addActionListener(new Export());
		frame.getContentPane().add(export);
		
		eredmeny = new ArrayList<String>();
	}
	
	//get only rows where cell values contain search string
		 
		static List<Row> getRows(XSSFSheet sheet, DataFormatter formatter, FormulaEvaluator evaluator, String searchValue) 								//list??z?? met??dus
		 {
			 List<Row> result = new ArrayList<Row>();													//??j t??mb l??trehoz??sa
			 String cellValue = "";																		//helyi string deklar??l??sa
			 
			 for (Row row : sheet) 																		//for each a t??bla bej??r??s??hoz
			 {
				 for (Cell cell : row) 																	//for each a sorok ??s cell??k bej??r??s??hoz
				 {
					 cellValue = formatter.formatCellValue(cell); 										//cell tartalm??nak odaad??sa egy v??ltoz??nak
					 
					 if (cellValue.equals(searchValue)) 												//cella tartalm??t ??sszehasonl??tja a keresett sz??val
					 {
						 result.add(row);																//result v??ltoz?? megkapja a cella tartalm??t ha volt egyez??s
						 //break;
					 }
				 }
			 }
		  return result;																				//visszat??r a result v??ltoz??
		 }
	
	class Kereses implements ActionListener																//keres??s gomb lenyom??sakor t??rt??nik			
	{
		 @SuppressWarnings("deprecation")
		public void actionPerformed(ActionEvent e)
		 {
			 
			model.removeAllElements();																	//t??rli a lista elemeit 
			String keresettszo = keresomezo.getText();													//keres??mez?? tartalm??t odaadja egy string v??ltoz??nak
			
			try
			{
				f = new File("z:\\RoHS,Reach, CFSI\\!CMRT\\");				//mappa beolvas??sa
				
				FilenameFilter filter = new FilenameFilter() 											//f??jln??v filter met??dus
				{
					
	                @Override
	                public boolean accept(File f, String name) 
	                {
	                    																				// csak az xlsx f??jlokat list??zza ki 
	                	return name.endsWith(".xlsx");	
	                }
	            };
	            
	            files = f.listFiles(filter);															//a beolvasott adatok egy f??jl t??mbbe rakja
	            
	            szurtfajlok = new ArrayList<String>();													//t??mlista deklar??l??sa
	            
	            /*
	            for (i = 0; i < files.length; i++)														//csak a megadott v??g?? elemeket list??zza ki a sz??kitett list??b??l
	            {
		            if(files[i].getName().indexOf("6.1") > -1) 
		            {
		            	szurtfajlok.add(files[i].getName());
		            	
		            }
		            
		            if(files[i].getName().indexOf("6.01") > -1)
		            {
		            	szurtfajlok.add(files[i].getName());
		            	
		            }
	            }
	            */
	            for (i = 0; i < files.length; i++)
	            {	
	            	
	            	FileInputStream fis = new FileInputStream("z:\\RoHS,Reach, CFSI\\!CMRT\\" +files[i].getName());															//f??jlok beolva??sa egyes??vel
				  XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis);  /*new XSSFWorkbook(fis); */															//excel f??jl t??rol?? l??trehoz??sa
				  //Workbook workbook = WorkbookFactory.create(new FileInputStream("./inputFile.xls"));
				  DataFormatter formatter = new DataFormatter();
				  FormulaEvaluator evaluator =  workbook.getCreationHelper().createFormulaEvaluator();
				  XSSFSheet sheet = workbook.getSheetAt(Integer.parseInt(tabla_szam.getText()));																			//a megadott sz??m?? t??bl??t t??rolja el
		
					  List<Row> filteredRows = keresoablak.getRows(sheet, formatter, evaluator, keresettszo);
		
						  for (Row row : filteredRows) 
						  {
						   for (Cell cell : row) 
						   {
						    System.out.print(cell.getAddress()+ ":" + formatter.formatCellValue(cell));																		//konzolra ??rja a tal??lt cella tartalm??t ??s sz??m??t
						    //System.out.print(" ");
						    model.addElement("Cella: " + cell.getAddress() + " Tartalma: " + formatter.formatCellValue(cell) + "  F??jl neve: "+ files[i].getName());		//Lit??hoz adja az eredm??nyt, cella tartalm??val, f??jl nev??vel 
						    eredmeny.add(cell.getAddress()+": "+ formatter.formatCellValue(cell) +"    " + files[i].getName());												//odaadja az eredm??nyt egy k??ztes t??mbnek, ami majd ki lesz ??rva egy excel f??jlba
						    
						    break;
						   }
						   System.out.println();
						  }
		
					  workbook.close();
					 
	            }
			
			}		
			catch (EncryptedDocumentException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);
			} 
			
			catch (FileNotFoundException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);
			} 
			catch (IOException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);
			}
			catch (java.lang.IllegalArgumentException e1) 
			{
				// TODO Auto-generated catch block
				
				JOptionPane.showMessageDialog(null, "Nincs ilyen t??bla sz??m az egyik excelben. F??jl neve: "+files[i].getName(), "T??j??koztat?? ??zenet", 1);		//hiba??zenet ablak megnyit??sa ha nem volt tal??lat
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);
			}
			catch (RuntimeException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);																	//hiba??zenet ablak megnyit??sa ha nem volt tal??lat
			} catch (InvalidFormatException e1) {
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);
			} 
			
			JOptionPane.showMessageDialog(null, "Keres??s v??get ??rt", "T??j??koztat?? ??zenet", 1);																	//keres??s v??g??n felugr?? ablak
		 }
	}
	
	class Export implements ActionListener																														//export gomb lenyom??sakor t??rt??nik			
	{
		public void actionPerformed(ActionEvent e)
		 {
			try
            {    
				XSSFWorkbook workbook = new XSSFWorkbook(); 																									//excel tipus?? v??ltoz?? l??trehoz??sa
	              
		        XSSFSheet sheet = workbook.createSheet("Eredm??nyek");																							//t??bla l??trehoz??sa a megadott n??vvel
		       
		        //Row row = sheet.createRow(0);
		        //Cell cell;
	            //cell = row.createCell(0);
	            int rownum = 0;
	            
		        for(String kereses: eredmeny)																													//for each a t??bla kiirat??s??hoz
		        {	
		        	Row row = sheet.createRow(rownum++);
		        	Cell cell = row.createCell(0);
		            cell.setCellValue(kereses);
		        }
				
				//sheet.getCellComments()).importArrayList(eredmeny, 0, 0, true);   
	            
	           
      			
				FileOutputStream out = new FileOutputStream(new File("z:\\RoHS,Reach, CFSI\\!CMRT smelter keres??sek\\" + keresomezo.getText() + " Smelter_sz??m.xlsx")); // f??jl tipus?? v??ltoz?? l??rtehoz??sa
				workbook.write(out);																															//ki??rja az el??bb megadott f??jlba az excel v??ltoz??t
				out.close();																																	//f??jl lez??r??sa
          
            } 
			catch (Exception e2) 
			{
				e2.printStackTrace();
			}
			
			JOptionPane.showMessageDialog(null, "Export??l??s v??get ??rt", "T??j??koztat?? ??zenet", 1);		//export??l??s v??g??n megjelen?? ??zenet
		 }
		
	}
	
}

