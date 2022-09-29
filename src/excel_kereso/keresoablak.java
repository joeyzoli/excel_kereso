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
					keresoablak window = new keresoablak();						//ablak létrehozása
					window.frame.setVisible(true);								//ablak láthatóvá tétele
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
		frame.setTitle("Kereső ablak");
		
		JButton keresogomb = new JButton("Keresés");
		keresogomb.setBounds(445, 63, 89, 23);
		keresogomb.addActionListener(new Kereses());
		frame.getContentPane().add(keresogomb);
		
		model = new DefaultListModel<String>();											//lista modell léterehozása
		lista = new JList<String>();
		lista.setBounds(127, 167, 407, 186);
		lista.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);					//1x-es kijelõlés beállítása
		lista.setModel(model);															//listamodell beállítása
		frame.getContentPane().add(lista);
		
		JScrollPane scrollPane = new JScrollPane(lista);								//scrollozható ablak létrehozása a Jlistből
		scrollPane.setBounds(127, 167, 407, 186);
		frame.getContentPane().add(scrollPane);
		
		keresomezo = new JTextField();
		keresomezo.setBounds(127, 64, 267, 20);
		frame.getContentPane().add(keresomezo);
		keresomezo.setColumns(10);
		
		JLabel keresettszo = new JLabel("Keresendő szó:");
		keresettszo.setBounds(127, 19, 131, 34);
		frame.getContentPane().add(keresettszo);
		
		JLabel lblNewLabel = new JLabel("Találatok:");
		lblNewLabel.setBounds(127, 142, 81, 14);
		frame.getContentPane().add(lblNewLabel);
		
		JLabel tabla_szama = new JLabel("Tábla száma");
		tabla_szama.setBounds(127, 95, 95, 14);
		frame.getContentPane().add(tabla_szama);
		
		tabla_szam = new JTextField();
		tabla_szam.setText(tablaszam);
		tabla_szam.setBounds(217, 92, 33, 20);
		frame.getContentPane().add(tabla_szam);
		tabla_szam.setColumns(10);
		
		JButton export = new JButton("Exportálás");
		export.setBounds(445, 389, 89, 23);
		export.addActionListener(new Export());
		frame.getContentPane().add(export);
		
		eredmeny = new ArrayList<String>();
	}
	
	//get only rows where cell values contain search string
		 
		static List<Row> getRows(XSSFSheet sheet, DataFormatter formatter, FormulaEvaluator evaluator, String searchValue) 								//listázó metódus
		 {
			 List<Row> result = new ArrayList<Row>();													//új tömb létrehozása
			 String cellValue = "";																		//helyi string deklarálása
			 
			 for (Row row : sheet) 																		//for each a tábla bejárásához
			 {
				 for (Cell cell : row) 																	//for each a sorok és cellák bejárásához
				 {
					 cellValue = formatter.formatCellValue(cell); 										//cell tartalmának odaadása egy változónak
					 
					 if (cellValue.equals(searchValue)) 												//cella tartalmát összehasonlítja a keresett szóval
					 {
						 result.add(row);																//result változó megkapja a cella tartalmát ha volt egyezés
						 //break;
					 }
				 }
			 }
		  return result;																				//visszatér a result változó
		 }
	
	class Kereses implements ActionListener																//keresés gomb lenyomásakor történik			
	{
		 @SuppressWarnings("deprecation")
		public void actionPerformed(ActionEvent e)
		 {
			 
			model.removeAllElements();																	//törli a lista elemeit 
			eredmeny.clear();
			String keresettszo = keresomezo.getText();													//keresőmező tartalmát odaadja egy string változónak
			
			try
			{
				f = new File("z:\\RoHS,Reach, CFSI\\!CMRT\\");											//mappa beolvasása
				
				FilenameFilter filter = new FilenameFilter() 											//fájlnév filter metódus
				{
					
	                @Override
	                public boolean accept(File f, String name) 
	                {
	                    																				// csak az xlsx fájlokat listázza ki 
	                	return name.endsWith(".xlsx");	
	                }
	            };
	            
	            files = f.listFiles(filter);															//a beolvasott adatok egy fájl tömbbe rakja
	            
	            szurtfajlok = new ArrayList<String>();													//tömlista deklarálása
	            
	            /*
	            for (i = 0; i < files.length; i++)														//csak a megadott végű elemeket listázza ki a szűkitett listából
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
	            	
	            	FileInputStream fis = new FileInputStream("z:\\RoHS,Reach, CFSI\\!CMRT\\" +files[i].getName());															//fájlok beolvaása egyesével
				  XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis);  /*new XSSFWorkbook(fis); */															//excel fájl tároló létrehozása
				  //Workbook workbook = WorkbookFactory.create(new FileInputStream("./inputFile.xls"));
				  DataFormatter formatter = new DataFormatter();
				  FormulaEvaluator evaluator =  workbook.getCreationHelper().createFormulaEvaluator();
				  XSSFSheet sheet = workbook.getSheetAt(Integer.parseInt(tabla_szam.getText()));																			//a megadott számú táblát tárolja el
		
					  List<Row> filteredRows = keresoablak.getRows(sheet, formatter, evaluator, keresettszo);
		
						  for (Row row : filteredRows) 
						  {
						   for (Cell cell : row) 
						   {
						    System.out.print(cell.getAddress()+ ":" + formatter.formatCellValue(cell));																		//konzolra írja a talált cella tartalmát és számát
						    //System.out.print(" ");
						    model.addElement("Cella: " + cell.getAddress() + " Tartalma: " + formatter.formatCellValue(cell) + "  Fájl neve: "+ files[i].getName());		//Litához adja az eredményt, cella tartalmával, fájl nevével 
						    eredmeny.add(cell.getAddress()+": "+ formatter.formatCellValue(cell) +"    " + files[i].getName());												//odaadja az eredményt egy köztes tömbnek, ami majd ki lesz írva egy excel fájlba
						    
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
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
			} 
			
			catch (FileNotFoundException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
			} 
			catch (IOException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
			}
			catch (java.lang.IllegalArgumentException e1) 
			{
				// TODO Auto-generated catch block
				
				JOptionPane.showMessageDialog(null, "Nincs ilyen tábla szám az egyik excelben. Fájl neve: "+files[i].getName(), "Tájékoztató Üzenet", 1);		//hibaüzenet ablak megnyitása ha nem volt találat
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
			}
			catch (RuntimeException e1) 
			{
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);																				//hibaüzenet ablak megnyitása ha nem volt találat
			} catch (InvalidFormatException e1) {
				// TODO Auto-generated catch block
				String hibauzenet = e1.toString();  
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
			} 
			
			JOptionPane.showMessageDialog(null, "Keresés véget ért", "Tájékoztató Üzenet", 1);																	//keresés végén felugró ablak
		 }
	}
	
	class Export implements ActionListener																														//export gomb lenyomásakor történik			
	{
		public void actionPerformed(ActionEvent e)
		 {
			try
            {    
				XSSFWorkbook workbook = new XSSFWorkbook(); 																									//excel tipusú változó létrehozása
	              
		        XSSFSheet sheet = workbook.createSheet("Eredmények");																							//tábla létrehozása a megadott névvel
		       
		        //Row row = sheet.createRow(0);
		        //Cell cell;
	            //cell = row.createCell(0);
	            int rownum = 0;
	            
		        for(String kereses: eredmeny)																													//for each a tábla kiiratásához
		        {	
		        	Row row = sheet.createRow(rownum++);
		        	Cell cell = row.createCell(0);
		            cell.setCellValue(kereses);
		        }
				
				//sheet.getCellComments()).importArrayList(eredmeny, 0, 0, true);   
	            
	           
      			
				FileOutputStream out = new FileOutputStream(new File("z:\\RoHS,Reach, CFSI\\!CMRT smelter keresések\\" + keresomezo.getText() + " Smelter_szám.xlsx")); 	// fájl tipusú változó lértehozása
				workbook.write(out);																																		//kiírja az elöbb megadott fájlba az excel változót
				out.close();																																				//fájl lezárása
          
            } 
			catch (Exception e2) 
			{
				e2.printStackTrace();
			}
			
			JOptionPane.showMessageDialog(null, "Exportálás véget ért", "Tájékoztató Üzenet", 1);																	//exportálás végén megjelenő üzenet
		 }
		
	}
	
}

