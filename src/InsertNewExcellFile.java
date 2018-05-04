import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.filechooser.FileSystemView;
import javax.xml.crypto.dsig.spec.HMACParameterSpec;

import java.awt.FlowLayout;
import java.awt.Window;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jdesktop.swingx.JXDatePicker;
import org.jdesktop.swingx.JXImagePanel.Style;

public class InsertNewExcellFile {

	public static void readXLSFile(Stoixeia a) throws IOException
	{
		Arxikopoihshegrafou nato=new Arxikopoihshegrafou();
		nato.setTitlos("Συγκεντρωτική κατασταση απουσιών");
		Date arxikh=a.getArxh();
		Date telikh=a.getTelos();
		System.out.println(arxikh);
		System.out.println(telikh);
		File f=new File(a.getPath());
		f.setWritable(true);
		InputStream ExcelFileToRead = new FileInputStream(a.getPath());
		HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		HSSFSheet sheet=wb.getSheetAt(0);
		int argrammwn=sheet.getLastRowNum();
		HSSFRow row; 
		HSSFCell cell;
		List<Student> st = new ArrayList();
	
		String titlossxol=sheet.getRow(4).getCell(0).getStringCellValue();
		System.out.println();
		//taksh
		String taksh=sheet.getRow(13).getCell(4).getStringCellValue();
		nato.setTaksh(taksh);
		System.out.println("taksh"+taksh);
		//sxoleio
		String sxoleio=sheet.getRow(3).getCell(0).getStringCellValue();
		System.out.println("sxoleio"+sxoleio);
		nato.setSxoleio(sxoleio);
		//
		String sxolikoetos=sheet.getRow(8).getCell(12).getStringCellValue();
		System.out.println(sxolikoetos);
		nato.setEtos(sxolikoetos);
		//
		String hmeromhniakatastashs=sheet.getRow(9).getCell(12).getStringCellValue();
		System.out.println(hmeromhniakatastashs);
		nato.setHmero(hmeromhniakatastashs);
		//diagrafh grammwn apo mhden ews 13
		for(int k=0;k<14;k++) {
			HSSFRow removingRow=sheet.getRow(k);
			sheet.removeRow(removingRow);
		}
		Iterator rows = sheet.rowIterator();
		Student neos;
		int metrhthsapoysiwn=0;
		int metrhthsapoarxh=0;
		//prin mpw edw na diagrapsw tis grammes poy prepei
		System.out.println("gamhmeno"+sheet.getRow(46).getCell(0).getCellType()+sheet.getRow(46).getCell(0).getNumericCellValue());
		while (rows.hasNext())
		{
			
			row=(HSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			
			while (cells.hasNext())
			{
				cell=(HSSFCell) cells.next();
				
				
				if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
				{
					if ((cell.getNumericCellValue() != HSSFCell.CELL_TYPE_BLANK || cell.getNumericCellValue()== 3.0 )&& cell.getColumnIndex()== 0 ) {
						neos= new Student(row.getCell(1).getStringCellValue(), row.getCell(2).getStringCellValue(), 0,0,0);
						System.out.println("arithmos grammhs pou to vrika"+row.getRowNum());
						System.out.println(neos.getEpwnimo()+cell.getNumericCellValue());
						st.add(neos);
					}
				}
				
				if (cell.getCellType()== HSSFCell.CELL_TYPE_STRING)
				{
					if (cell.getStringCellValue() == sheet.getRow(argrammwn).getCell(1).getStringCellValue() /*&& cell.getColumnIndex()== 1*/) {
						int size=st.size();
						neos=st.get(size-1);
						neos.setApousies((int) row.getCell(5).getNumericCellValue());
						neos.setMerikesapousies(metrhthsapoysiwn);
						neos.setApousiesApoArxh(metrhthsapoarxh);
						st.remove(size-1);
						st.add(neos);
						//mhdenise ton metrhth apousiwn
						metrhthsapoysiwn=0;
						metrhthsapoarxh=0;
				}
				}
			
				if (cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC) {
					if (cell.getColumnIndex()== 3 ) {
						System.out.println(cell.getDateCellValue());
						if((cell.getDateCellValue().after(arxikh) || cell.getDateCellValue().equals(arxikh)  ) && (cell.getDateCellValue().before(telikh)|| cell.getDateCellValue().equals(telikh))) {
							metrhthsapoysiwn+=row.getCell(5).getNumericCellValue();
							
						}
						if(cell.getDateCellValue().before(telikh)|| cell.getDateCellValue().equals(telikh)) {
							metrhthsapoarxh+=row.getCell(5).getNumericCellValue();
						}
					}
				}
				
				
			
			}
		
		}
		System.out.println("Επωνυμο Ονομα       Συνολο Απουσιών               μερικες απουσιες");
		for (Student student : st) {
			System.out.println(student.getEpwnimo()+ " "+student.getOnoma() + "       "+ student.getApousies()+"         "+student.getMerikesapousies());
		}
		System.out.println(st.size() + "mhkos alusidas");
	writeXLSFile(st,a,nato);
	if (a.getF2()!=null) {
		write2XLSFile(st,a,nato);
	}
	}
	
	
	
	
	private static void write2XLSFile(List<Student> st, Stoixeia a, Arxikopoihshegrafou nato) throws IOException{
		File f=new File(a.getPath2());
		f.setWritable(true);
		//thelei prwta read kai meta write se kainourgio 
		//diavasma
		
		FileInputStream fos=new FileInputStream(a.getPath2());
		HSSFWorkbook wb = new HSSFWorkbook(fos);
		HSSFSheet sheet = wb.getSheetAt(0) ;
		HSSFRow row; 
		HSSFCell cell;
		List<Emails> m = new ArrayList();
		HSSFCellStyle arx=sheet.getRow(0).getCell(0).getCellStyle();
		HSSFRow removingRow=sheet.getRow(0);
		sheet.removeRow(removingRow);
		Iterator rows = sheet.rowIterator();	
		
		while (rows.hasNext())
		{
			row=(HSSFRow) rows.next();
			Emails ems=new Emails();
			//karfwta

			ems.setMath(row.getCell(1).getStringCellValue());
			ems.setDiastma(row.getCell(2).getStringCellValue());
			ems.setProsma(row.getCell(5).getStringCellValue());
			ems.setProsgonea(row.getCell(6).getStringCellValue());
			ems.setGoneas(row.getCell(7).getStringCellValue());
			ems.setMail(row.getCell(8).getStringCellValue());
			ems.setKin((long)row.getCell(9).getNumericCellValue());
			m.add(ems);
		}
		
		
		
		
		//grapsimo
		//meion ena giati pairnw kai thn prwth grammh
		String excelFileNameWrite = a.getF2()+File.separator+"sg_"+a.getFname2();//name of excel file
		HSSFWorkbook wb1 = new HSSFWorkbook();
		String sheetName = "Sheet1";
		HSSFSheet sheet1 = wb1.createSheet(sheetName) ;
		SimpleDateFormat eDateFormat= new SimpleDateFormat("dd/MM/yyyy");
		String apoews=eDateFormat.format(a.getArxh())+"-"+eDateFormat.format(a.getTelos());
		HSSFRow rowar6 = sheet1.createRow(0);
		HSSFCell cell100 = rowar6.createCell(0);
		cell100.setCellValue("ΑΑ");
		/*HSSFCellStyle h=cell100.getCellStyle();
		h.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		h.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		cell100.setCellStyle(h);
		*/
		HSSFCell cell101 = rowar6.createCell(1);
		cell101.setCellValue("Στοιχεια Μαθητη");
		HSSFCell cell102 = rowar6.createCell(2);
		cell102.setCellValue("Διαστημα");
		HSSFCell cell103 = rowar6.createCell(3);
		cell103.setCellValue("ΑΠΟΥΣ_ ΔΙΑΣΤ" );
		HSSFCell cell104 = rowar6.createCell(4);
		cell104.setCellValue("ΣΥΝ_ΑΠΟΥΣΙΩ");
		HSSFCell cell105 = rowar6.createCell(5);
		cell105.setCellValue("ΠΡΟΣ_ΜΑΘ");
		HSSFCell cell106 = rowar6.createCell(6);
		cell106.setCellValue("ΠΡΟΣ_ΚΗΔΕΜΟΝΑ");
		HSSFCell cell107 = rowar6.createCell(7);
		cell107.setCellValue("ΚΗΔΕΜΟΝΑΣ");
		HSSFCell cell108 = rowar6.createCell(8);
		cell108.setCellValue("MAIL");
		HSSFCell cell109 = rowar6.createCell(9);
		cell109.setCellValue("ΚΙΝΗΤΟ ΤΗΛ");
	
		
		
		
		
		
		for(int i=0;i<st.size();i++) {
			HSSFRow row1=sheet1.createRow(i+1);
			for(int c=0;c<10;c++) {
				HSSFCell cell1 = row1.createCell(c);
			    switch (c) {
				case 0:
					cell1.setCellValue(i+1);
					break;

				case 1:
					cell1.setCellValue(m.get(i).getMath());
					break;
				case 2:
					cell1.setCellValue(apoews);
					break;
				case 3:
					cell1.setCellValue(st.get(i).getApousiesApoArxh());
					break;
				case 4:
					cell1.setCellValue(st.get(i).getMerikesapousies());
					break;
				case 5:
					cell1.setCellValue(m.get(i).getProsma());
					break;
				case 6:
					cell1.setCellValue(m.get(i).getProsgonea());
					break;
				case 7:
					cell1.setCellValue(m.get(i).getGoneas());
					break;
				case 8:
					cell1.setCellValue(m.get(i).getMail());
					break;
				case 9:
					cell1.setCellValue(m.get(i).getKin());
					break;
					
				default:
					break;
				}
			}
		}
	for(int i=0;i<10;i++) {
		sheet1.autoSizeColumn(i);
	}
	FileOutputStream foska=new FileOutputStream(excelFileNameWrite);
	wb1.write(foska);
	foska.flush();
	foska.close();
	
	}




	public static void writeXLSFile(List <Student> st,Stoixeia pr,Arxikopoihshegrafou nato) throws IOException {
		String Filepath=FileSystemView.getFileSystemView().getHomeDirectory().toString();
		String excelFileNameWrite = pr.getF()+File.separator+"sg_"+pr.getFname();//name of excel file
		String sheetName = "Sheet1";//name of sheet
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet(sheetName) ;
		 {
			 CellRangeAddress cellRangeAddress=new CellRangeAddress(0, 0, 0, 4);
			 sheet.addMergedRegion(cellRangeAddress);
			 HSSFRow rowar0=sheet.createRow(0);
			 HSSFCell cell00=rowar0.createCell(0);
			 cell00.setCellValue(nato.getSxoleio());
			 
			 
			 
			 //αρχικοποιηση πρώτησ γραμμης me titlo
			 CellRangeAddress c1=new CellRangeAddress(1, 1, 0, 4);
			sheet.addMergedRegion(c1);	 
			 HSSFRow rowar1=sheet.createRow(1);
				 HSSFCell cell10=rowar1.createCell(0);
				 cell10.setCellValue(nato.getTitlos());
				
			 
			 //apo 
				 CellRangeAddress c3=new CellRangeAddress(2, 2, 1, 3);
				 sheet.addMergedRegion(c3);
				 HSSFRow rowar2=sheet.createRow(2);
				 HSSFCell cell20=rowar2.createCell(0);
				 cell20.setCellValue("ΑΠΟ ΕΩΣ");
				 HSSFCell cell21=rowar2.createCell(1);
				 SimpleDateFormat eDateFormat= new SimpleDateFormat("dd/MM/yyyy");
				 cell21.setCellValue(eDateFormat.format(pr.getArxh())+"-"+eDateFormat.format(pr.getTelos()));
				  
			
			 //ews
			/*
				 HSSFRow rowar3=sheet.createRow(3);
				 HSSFCell cell30=rowar3.createCell(0);
				 cell30.setCellValue("ΜΕΧΡΙ");
				 HSSFCell cell31=rowar3.createCell(1);
				 cell31.setCellValue(eDateFormat.format(pr.getTelos()));
			*/
			 //tasksh
			
				 HSSFRow rowar4=sheet.createRow(3);
				 HSSFCell cell40=rowar4.createCell(0);
				 cell40.setCellValue("ΤΑΞΗ");
				 HSSFCell cell41=rowar4.createCell(1);
				 cell41.setCellValue(nato.getTaksh());
			
				 CellRangeAddress c2=new CellRangeAddress(4, 4, 0,1);
				 sheet.addMergedRegion(c2);
				 HSSFRow rowar5=sheet.createRow(4);
				 HSSFCell cell50=rowar5.createCell(0);
				 cell50.setCellValue("ΗΜ/ΝΙΑ ΛΗΨΗΣ ΑΡΧΕΙΟΥ");
				 HSSFCell cell51=rowar5.createCell(2);
				 cell51.setCellValue(nato.getHmero());
	
				HSSFRow rowar6 = sheet.createRow(10);
				HSSFCell cell100 = rowar6.createCell(0);
				cell100.setCellValue("ΑΑ");
				HSSFCell cell101 = rowar6.createCell(1);
				cell101.setCellValue("Επωνυμο");
				HSSFCell cell102 = rowar6.createCell(2);
				cell102.setCellValue("Ονομα");
				HSSFCell cell103 = rowar6.createCell(3);
				cell103.setCellValue("ΑΠ_ΔΙΑΣΤ");
				HSSFCell cell104 = rowar6.createCell(4);
				cell104.setCellValue("ΕΩΣ ΚΑΙ");
				HSSFCell cell105 = rowar6.createCell(5);
				cell105.setCellValue("ΣΥΝ_ΑΠ");
				
			
		 }
		
		//iterating r number of rows
		for (int r=0;r < st.size(); r++ )
		{
			HSSFRow row = sheet.createRow(r+11);
				for(int c=0;c<6;c++) {
					HSSFCell cell = row.createCell(c);
					switch (c) {
						case 0:
							cell.setCellValue(r+1);
							break;
						case 1:
							cell.setCellValue(st.get(r).getEpwnimo());
							break;
						case 2:
							cell.setCellValue(st.get(r).getOnoma());
							break;
						case 3:
							if(st.get(r).getMerikesapousies()==0) {
								cell.setCellValue(" ");
							}else {
							cell.setCellValue(st.get(r).getMerikesapousies());		
							}
							break;
						case 4:
							cell.setCellValue(st.get(r).getApousiesApoArxh());
							break;
						case 5:
								cell.setCellValue(st.get(r).getApousies());
								break;
							
					}		
				}
				
			}
		
		

		 sheet.autoSizeColumn(0);
		 sheet.autoSizeColumn(1);
		 sheet.autoSizeColumn(2);
		 sheet.autoSizeColumn(3);
		 sheet.autoSizeColumn(4);
		
		FileOutputStream fileOut = new FileOutputStream(excelFileNameWrite);
		
		//write this workbook to an Outputstream.
		wb.write(fileOut);
		
		 
		fileOut.flush();
		fileOut.close();
	}
	
	public static void main(String[] args) throws IOException {
	    Stoixeia facts=new Stoixeia();
		JFrame frame = new JFrame("Μετατροπη απο το School.gr σε αρχειο excell ");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(300, 280);
		JPanel controlpanel = new JPanel();
		controlpanel.setLayout(new FlowLayout());
		JLabel arxh=new JLabel("Ημερομηνια αρχής");
		JXDatePicker datepickerarxh=new JXDatePicker();
		SimpleDateFormat formats=new SimpleDateFormat("dd/MM/yyyy");
		datepickerarxh.setFormats(formats);
		JLabel telos=new JLabel("Ημερομηνια λήξης");
		JTextArea a=new JTextArea();
		JTextArea b=new JTextArea();
		JXDatePicker datepickertelos=new JXDatePicker();
		datepickertelos.setFormats(formats);
		JButton kleise= new JButton("Κλείσιμο");
		JLabel mnm = new JLabel("Συμπληρωσε το αρχειο και τις ημερομηνιες");
		datepickerarxh.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				System.out.println(datepickerarxh.getDate());
				facts.setArxh(datepickerarxh.getDate());
			}
		} );
		datepickertelos.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				System.out.println(datepickertelos.getDate());
				facts.setTelos(datepickertelos.getDate());
			}
		});
		kleise.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				System.exit(0);
			}
		});
		
		JButton epeks= new JButton("Διαλεξέ το αρχείο προς επεξεργασία");
		epeks.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				JFileChooser fc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
				int rtvalue= fc.showOpenDialog(null);
				if (rtvalue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fc.getSelectedFile();
					System.out.println(selectedFile.getAbsolutePath());
					System.out.println(selectedFile.getParentFile());
					System.out.println(selectedFile.getName());
					String pathfile = selectedFile.getAbsolutePath();
					a.setText(pathfile);
					facts.setF(selectedFile.getParentFile());
					facts.setPath(selectedFile.getAbsolutePath());
					facts.setFname(selectedFile.getName());
					if(facts!=null) {
						if(facts.getArxh()!=null && facts.getTelos()!=null) {
						try {
							readXLSFile(facts);
							b.setText(selectedFile.getParentFile()+File.separator+"sg"+selectedFile.getName());
						}catch(IOException e){
							e.printStackTrace();
						}
						}else {
					JOptionPane.showMessageDialog(frame, "Πρέπει πρώτα να επιλέξετε ημερομηνία αρχής και λήξης και έπειτα να επιλέξετε αρχείο εισόδου.");
						
						}
					}
				}
				
			}
		});
		JButton em= new JButton("Διαλεξέ το αρχείο me ta emails");
		em.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser fc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
				int rtvalue= fc.showOpenDialog(null);
				if (rtvalue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fc.getSelectedFile();
					facts.setF2(selectedFile.getParentFile());
					facts.setPath2(selectedFile.getAbsolutePath());
					facts.setFname2(selectedFile.getName());
				}
			}
		});
		controlpanel.add(mnm);
		controlpanel.add(arxh);
		controlpanel.add(datepickerarxh);
		controlpanel.add(telos);
		controlpanel.add(datepickertelos);
		controlpanel.add(em);
		controlpanel.add(epeks);
		controlpanel.add(kleise);
		controlpanel.add(a);
		controlpanel.add(b);
		frame.add(controlpanel);
		frame.setVisible(true);
		
		
		
	}
	
}
