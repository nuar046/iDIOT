package my.org;

import java.sql.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFConditionalFormattingRule;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSheetConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.AreaPtgBase;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;

/** GPL v2.0 LICENSE. COPYRIGHTS INHERENT AT LEAST BUT NOT LIMITED TO COMPIERE, APACHE, ITEXTPDF ICEBLUE projects
 *  PLS INFORM THE AUTHOR EMAIL BELOW OF ANY INGRINGEMENT AND THE CODE SHALL BE EXPUNGED FROM LATEST VERSION
 *  @version 1.0a
 *  Any novelty within here is Copyright (C) 2022 Redhuan D. Oon and stated contributors
 *  DISCLAIMER  - THIS PROJECT IS A FLOSS EXPERIMENT NOT FOR COMMERCIAL INTENT. IT IS MERELY TO PROVE THE AUTHORS NOVELTY
 * 
 * THIS REFERS TO FACTACCTREPORT.XLS PLS REFER NOTES THERE
 * SHEETS: INPUT, PROCESS, OUTPUT 
 * INPUT IS ACTUAL LAYOUT WITH ANNOTATIONS
 * OUTPUT IS WELL DESIGNED COLORED LAYOUT TO BE WRITTEN
 * PROCESS SHEET SAMPLE SHOWN BELOW AS CONTROL FOR OUTPUT
	
	ANNOTATION			ROW	SELECT			TABLE	 	WHERE  									ADDRESS		Direction	VALUES
	@a.Fact_Acct_ID@2	2	a.Fact_Acct_ID	Fact_Acct a	(a.datetrx BETWEEN #1 and  #2 )				D7		>			1008011
	@a.AmtAcctDr@3		3	a.AmtAcctDr					#?a.DateTrx = [C8] AND c.DocumentNo = [B7]	E8		S			200,000 
	@c.DocumentNo@4		4	c.DocumentNo	LEFT JOIN v_docno c on c.table_id=a.ad_table_id			B7		V			1000021
 * 
 * @author red1org@gmail.com Redhuan D. Oon for MOTIVE SOLUTIONS THAILAND
 * and ahmad.anwar.ibrahim@gmail.com and contributors
 */
public class idiot {
	    static String DB_URL = null;
	    static String USER = null;
	    static String PASS = null;
	    private static String File_Directory = "C:\\Users\\60133\\Downloads\\dubizzle.xls"; 
		Timestamp DateFrom = null;
		Timestamp DateTo = null;
		private int AD_Org_ID = 0;
		private static String AccountGroup = "";
		static HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
		static HSSFWorkbook workbook = new HSSFWorkbook();   
		static String direction = "";
		final static int vcol = 8;
		static int sqlcount = 0;
		static String DOWN = "V";
		static String LEFT = "<";
		static String RIGHT = ">";
		static String END = "END";
		static String PLUS = "\\+";
		static int lastrowwrite = 1;
		static int lastrowread = 0; 
		String documentno = "";
		static long start = System.currentTimeMillis(); 
		//SET ParameterTag values for #1 From, #2 To	
		static String from = "";
		static String to = "";
		static boolean gotSQL = true;
        static String path = FilenameUtils.getFullPath(File_Directory);
        static HSSFSheet configsheet = null;
        static HSSFSheet outputsheet = null;
		static HSSFSheet inputsheet = null;
		static HSSFSheet processsheet = null; 
		static boolean crosstab;   
		static HSSFSheetConditionalFormatting conditions = null;
		public static final String localepattern =  "#,##0.##;(#,##0.##)";
		static String File_Directory_PDF = "";
		static final String FONT = "font/NotoSerifLao/NotoSerifLao-Regular.ttf"; 
		
public static void main(String[] args) throws AlsCustomException, SQLException, InterruptedException {
		try {
		   System.out.println("Argument count: " + args.length);
		   for (int i = 0; i < args.length; i++) {
		      System.out.println("Argument " + i + ": " + args[i]);
		   }   
			sqlcount = 0;
			direction = RIGHT;
			FileInputStream file = new FileInputStream(File_Directory);    
			path = FilenameUtils.getFullPath(File_Directory);
			workbook = new HSSFWorkbook(file); 
			int norecs = 0;
			List<List<Object>> results = null;

			clean(file); //clean up write and close back
			
			workbook.setForceFormulaRecalculation(true); 
			configsheet = workbook.getSheet("Config");
			inputsheet = workbook.getSheet("Input");
			processsheet = workbook.getSheet("Process"); 
			outputsheet = workbook.getSheet("Output");
			setConfig(configsheet);			
			putImage(inputsheet,outputsheet);
			setAddress(inputsheet,processsheet,outputsheet); 
			//
			while (gotSQL) { //there is a running tracker lastrowread / lastrowwrite
				sqlcount++;
				StringBuilder completeStatement = makeFullSQLStatement(processsheet);
				//
				int br = completeStatement.indexOf("[");
				if (!crosstab && br>0)
					completeStatement = parseBracketTag(outputsheet, completeStatement);
				results = executeSQL(completeStatement);
				//
				if (results!=null){
					norecs+=results.size();
					set(results,processsheet);
					if (crosstab)
						crosstab(processsheet,outputsheet);
				}
				 lastrowwrite = lastrowread;
			}
			// 
			checkConditionalFormatting();
			paint(processsheet,outputsheet);
			System.out.println("END");
			workbookWrite();
			workbookClose(workbook);
			doPDF();
			//
			long finish = System.currentTimeMillis();
			long timeElapsed = finish - start; 
			BigDecimal divisor = new BigDecimal("100");
			System.out.println(norecs+" Records. Time Elapsed: "+BigDecimal.valueOf(timeElapsed)
			.divide(divisor).divide(BigDecimal.TEN)+" secs");
		} 
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} 
		catch (IOException e) {
			e.printStackTrace();
		} 
	}

private static void setConfig(HSSFSheet configsheet) {

    Row row = configsheet.getRow(1);
    Cell cell = row.getCell(1);
    DB_URL = cell.getStringCellValue();
    row = configsheet.getRow(2);
    cell = row.getCell(1);
    USER = cell.getStringCellValue();
    row = configsheet.getRow(3);
    cell = row.getCell(1);
    Double n = cell.getNumericCellValue(); 
    Integer temppass = n.intValue();
    PASS = temppass.toString();
}

/**
 * 
 * @throws IOException
 * @throws InterruptedException 
 * @throws DocumentException
 */ 
static void doPDF() throws IOException, InterruptedException{ 
		workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      	HSSFSheet pshit = workbook.getSheet("Output"); 
        // To iterate over the rows 
        File_Directory_PDF = File_Directory.replace("xls", "pdf");
        final PdfWriter pdfWriter = new PdfWriter(File_Directory_PDF);
        final PdfDocument pdfDoc = new PdfDocument(pdfWriter);
        Document doc = new Document(pdfDoc,PageSize.A3.rotate());
        PdfFont f = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);
  	    int lastrow = pshit.getLastRowNum();
  	    HSSFRow row = pshit.getRow(1);
  	    int lastcol = row.getLastCellNum();
  	    int hidden = 0;
  	    for (int i=0;i<lastcol;i++) {
  	    	if (outputsheet.isColumnHidden(i))
  	    		hidden++;
  	    }
        Table table = new Table(lastcol-hidden);
  	    table.setWidth(UnitValue.createPercentValue(100));   
  	    DecimalFormat df = (DecimalFormat) DecimalFormat.getInstance();
  	    for (int i=0;i<=lastrow;i++) {
  	    	row = pshit.getRow(i);
  	    	if (row==null)
  	    		continue;
  	    	if (row.getZeroHeight())
  	    		continue;
  	    	for (int c=0;c<lastcol;c++) {
  	    		if (outputsheet.isColumnHidden(c))
  	    			continue;
  	    		HSSFCell cell = row.getCell(c); 
	  	    	Paragraph p = new Paragraph().setMarginLeft(2f).setMarginRight(2f); 
	    		com.itextpdf.layout.element.Cell cellt = new com.itextpdf.layout.element.Cell();
	    		
	    		if (cell==null) {  
 					cellt.add(p).setBorder(Border.NO_BORDER);
					table.addCell(cellt); 
					continue;
  	    		}
	    		if (cell.getCellStyle().getBorderTop()==BorderStyle.NONE)
        			cellt.setBorderTop(Border.NO_BORDER);
        		if (cell.getCellStyle().getBorderRight()==BorderStyle.NONE)
        			cellt.setBorderRight(Border.NO_BORDER);
        		if (cell.getCellStyle().getBorderBottom()==BorderStyle.NONE)
					cellt.setBorderBottom(Border.NO_BORDER);
        		if (cell.getCellStyle().getBorderLeft()==BorderStyle.NONE)
        			cellt.setBorderLeft(Border.NO_BORDER);
  	    		if (conditions!=null)
  	    			if (pdfConditionalFormatting(cell)) {
  	    				p.add(" ");
						cellt.add(p);
						table.addCell(cellt); 
						continue;
  	    			} 

  	    		Short high = cell.getCellStyle().getFont(workbook).getFontHeight();
  	    		float fontsize = high/(6+(lastcol-hidden));
  	    		p.setFontSize(fontsize);
  	    		int div = (int)(9-(fontsize/(fontsize-2)));
		        if(cell.getCellType() == CellType.NUMERIC){ 
		        	if (DateUtil.isCellDateFormatted(cell)){
		        	DataFormatter formatter = new DataFormatter(); 
		        	p.add(formatter.formatCellValue(cell));
		        	}else {
			        	Double n = cell.getNumericCellValue(); 
			        	df.applyPattern(localepattern);
	                    p.add(df.format(n)).setTextAlignment(TextAlignment.RIGHT);
	                    }
                    cellt.add(p);
					table.addCell(cellt); 
		        	} 
		        else if(cell.getCellType() == CellType.STRING){ 
		        	HSSFCell cell2 = row.getCell(c+1); 
	        		int length = cell.getStringCellValue().length();
		        	int span = length/div;
		        	if (span>1) {
		        		if (cell2==null || cell2.getCellType()==CellType.BLANK) { 
		        		if (span>5) //limit to 5 cols
		        			span=6;
		        		cellt = new com.itextpdf.layout.element.Cell(0,span);
		        		c = c+span-1;
		        		if (cell.getCellStyle().getBorderTop()==BorderStyle.NONE)
		        			cellt.setBorderTop(Border.NO_BORDER);
		        		if (cell.getCellStyle().getBorderRight()==BorderStyle.NONE)
		        			cellt.setBorderRight(Border.NO_BORDER);
		        		if (cell.getCellStyle().getBorderBottom()==BorderStyle.NONE)
							cellt.setBorderBottom(Border.NO_BORDER);
		        		if (cell.getCellStyle().getBorderLeft()==BorderStyle.NONE)
		        			cellt.setBorderLeft(Border.NO_BORDER);
		        		}
		        	}	
		        	
					p.add(getunicode(cell.getStringCellValue()).setFont(f));	
					cellt.add(p);
					table.addCell(cellt);
	            	}  
	            else if(cell.getCellType() == CellType.FORMULA) { 
	            	switch(cell.getCachedFormulaResultType()) {
	                    case NUMERIC:
	                    	Double n = cell.getNumericCellValue(); 
	                        df.applyPattern(localepattern);
	                        p.add(df.format(n)).setTextAlignment(TextAlignment.RIGHT);
	    					cellt.add(p);
	    					table.addCell(cellt); 
	                        break;
	                    case STRING:
	                    	HSSFCell cell2 = row.getCell(c+1); 
	                    	int length = cell.getStringCellValue().length();
				        	int span = length/div;
				        	if (span>1) {
				        		if (cell2==null || cell2.getCellType()==CellType.BLANK) { 
					        		if (span>5)
					        			span=5;
					        		cellt = new com.itextpdf.layout.element.Cell(0,span);
					        		if (cell.getCellStyle().getBorderTop()==BorderStyle.NONE)
					        			cellt.setBorderTop(Border.NO_BORDER);
					        		if (cell.getCellStyle().getBorderRight()==BorderStyle.NONE)
					        			cellt.setBorderRight(Border.NO_BORDER);
					        		if (cell.getCellStyle().getBorderBottom()==BorderStyle.NONE)
										cellt.setBorderBottom(Border.NO_BORDER);
					        		if (cell.getCellStyle().getBorderLeft()==BorderStyle.NONE)
					        			cellt.setBorderLeft(Border.NO_BORDER);
					        		c = c+span-1;	
				        		}
				        	}
	                    	p.add(getunicode(cell.getStringCellValue()).setFont(f));	
	     					cellt.add(p);
	        				table.addCell(cellt);
	                        break;
					default:
						break;
	                }
	             } else if (cell.getCellType() == CellType.BLANK) {
	  	    			p.add(" ");
						cellt.add(p);
						table.addCell(cellt); 
	             }
  	    	}   
		}
  	  try {
  		String Org = "MOES";
		  
  	    String pathToFile = path+"/"+Org+".jpeg";
  	    File gotF = new File(pathToFile);
  	    if(gotF.exists()){
  	    	ImageData data = ImageDataFactory.create(pathToFile); 
	  	    Image img = new Image(data);
	  	    img.setFixedPosition(250, 750);
	  	    img.scaleToFit(44, 44);
	  	    doc.add(img);
  	    	}
	  	} catch (IOException ex) {
	  	    ex.printStackTrace();
	  	}
	        doc.add(table);
	        doc.close();
		}
	
	private static boolean pdfConditionalFormatting(HSSFCell cell) {
		if (cell==null || cell.getCellType()==CellType.BLANK)
			return false;
        int row = cell.getRowIndex();
        int col = cell.getColumnIndex();
  
       for (int idx = 0; idx < conditions.getNumConditionalFormattings(); idx++) {
        	HSSFConditionalFormatting cf = conditions.getConditionalFormattingAt(idx);
        	List<CellRangeAddress> cra = Arrays.asList(cf.getFormattingRanges());  
		    for (CellRangeAddress c : cra) { 
		    	if (c.isInRange(cell)) {
		    		HSSFConditionalFormattingRule rule = cf.getRule(0);
		    		if (rule==null)
		    			return false;
		    	     rule.getConditionType();
					if (ConditionType.FORMULA != null) {
		    	    	 String[] split = rule.getFormula1().split("=");
		    	    	 char ch =  split[0].charAt(split[0].length()-1);
		    	    	 CellReference ref = new CellReference(cell);
		    	    	 HSSFRow r = outputsheet.getRow(ref.getRow());
		    	    	 int pos = ch - 'A';
		    	    	 HSSFCell ce = r.getCell(pos);
		    	    	 String alpha = split[1].replaceAll("\"", "");
		    	    	 if (ce.getStringCellValue().equals(alpha)) 
		    	    		 return true;
		    	    	 else return false;
		    	     }
		    	}
		    }
        }
	return false;
}

	private static Text getunicode(String rawstring)
	{
		String unicodestring="";
		for (int i = 0; i < rawstring.length(); i++) 
		{
			if(rawstring.charAt(i) == ' ') {
				unicodestring = unicodestring + ' ';
			}
			else {
			unicodestring = unicodestring + "\\u" + Integer.toHexString(rawstring.charAt(i) | 0x10000).substring(1);
			}
		} 
		return new Text(StringEscapeUtils.unescapeJava(unicodestring));
	} 

private static void workbookClose(HSSFWorkbook wb) throws IOException { 
	wb.close();
}

private static void paint(HSSFSheet prsheet, HSSFSheet outsheet) throws AlsCustomException,IOException {
	//take column "Address" as ready array
	Iterator<Row> rowIteratePr = prsheet.rowIterator();   
		
	rowIteratePr.next();//row tracker starts at 1 and shall be running number
	
	while (rowIteratePr.hasNext()) {  
		int nc = vcol;//column VALUEs start at 7
		HSSFRow rowPr = (HSSFRow)rowIteratePr.next(); 
		//search address in insheet
		HSSFCell addressPr = rowPr.getCell(6); 
		if (rowPr.getCell(0)==null)
			break;
		if (addressPr==null || addressPr.toString().isBlank())
			continue;
		String annotation = rowPr.getCell(0).getStringCellValue()+rowPr.getCell(7).getStringCellValue(); 
		//fetch values from prSheet
		CellReference ref = new CellReference(addressPr.getStringCellValue());
		HSSFRow outrow = outsheet.getRow(ref.getRow());
		if (outrow==null)
			outrow = outsheet.createRow(ref.getRow());
		HSSFCell outcell = outrow.getCell(ref.getCol());
		if (outcell == null) 
			outcell = outrow.createCell(ref.getCol());  
		//
		//holder for row to go DOWN
		int downrow = outrow.getRowNum();
		HSSFRow outwriterow = outrow;
		
		//fetch/write VALUEs to outputsheet    
		boolean prValueNotEnd = true;
		boolean onetime = true;
		HSSFCell prValue = rowPr.getCell(nc);
		while (prValueNotEnd) {  
			if (outcell==null)
				break;
			if (nc>99)
				break;
			if (prValue == null)
				break;
			if (prValue.toString().equals(END))
				break;
			if (prValue.toString().equals("NULL")) {
				nc++;
				prValue = rowPr.getCell(nc);
				continue;
			}
			if (prValue.getCellType() == CellType.BLANK || prValue.getCellType()==CellType.FORMULA) {
				nc++;
				prValue = rowPr.getCell(nc);
				continue;
			}
			if (prValue.getCellType()==CellType.NUMERIC) { 
				   outcell.setCellValue(prValue.getNumericCellValue());							   
			   } else {
				   outcell.setCellValue(prValue.getStringCellValue());
			   }
			outcell.setCellStyle(prValue.getCellStyle());
			
			//get next value writing - 
			//there is RIGHT AND DOWN direction to display subsequent records
			if (annotation.endsWith(DOWN)){ //DOWNWARDS
				downrow++;
				outwriterow = outsheet.getRow(downrow);
				if (outwriterow==null)
					outwriterow = outsheet.createRow(downrow);
				int c = outcell.getColumnIndex();
				outcell = outwriterow.getCell(c);
				//
				if (outcell==null)
					outcell = outwriterow.createCell(c);
				
				//Check need to insert row by peek of next prValue exist
				outcell = pushRowDown(outsheet, nc, rowPr, outcell, downrow, c);
				
			} else if (annotation.endsWith(RIGHT)){ //continue RIGHT-WISE
				if (outrow.getCell(outcell.getColumnIndex()+1)==null)
					outcell = outrow.createCell(outcell.getColumnIndex()+1);
				else
					outcell = outrow.getCell(outcell.getColumnIndex()+1);
				//Check need to insert row by peek of next prValue exist
				HSSFCell peek = rowPr.getCell(nc+1);
				if (outcell.getCellType()==CellType.FORMULA && peek!=null && !peek.toString().equals(END)){
					if (!peek.toString().isBlank()) {
						// INSERT column 
						int atColumn = outcell.getColumnIndex()-1; //W-22 has formula to block
					    int endColumn = outrow.getLastCellNum();   //X-23 - next column after last
					    pushColumnsRight(outsheet,atColumn,endColumn); //last column at
					    outcell = outrow.getCell(atColumn+1);//guarantee the next column to write
					}			
				}
			} else if (annotation.contains("+")) {
				int v = vcol-2;
				int pv = vcol;
				outcell.setBlank(); //revert back as this is The Matrix
				Row mrow = outrow;
				Cell mcell = outcell;
				/**
				 * TODO GRID XTAB PIVOT TABLE
				 * 1. Check  'Addresses' instead of >,V i.e. N7+B13 
				 * 2. Address points to X an Y axes 
				 * 3. Take addressed value set for same indexes to form the fuller bigger matrix
				 * 4. Values placed according to matching X and Y indexes  
				 * 
				 */
				//is XTAB CHECK PIVOT TABLE
				String[] xtabcode = rowPr.getCell(7).getStringCellValue().split(PLUS);
				//find length of grid tabs
				CellReference refX = new CellReference(xtabcode[0]);
				Iterator<Row> rowIterateMatrix = prsheet.rowIterator();  
				rowIterateMatrix.next();//row tracker starts at 1 and shall be running number 
				ArrayList<Object> yaxis = new ArrayList<Object>();
				ArrayList<Object> xaxis = null;
				int xint = 0;
				int yint = 0;
				Object[][] matrix = null;//direct pr-values
				Object[][] canvas = null;//paint position
				HSSFRow rowX = null;
				HSSFCell cellX = null; 
				String dir = "";
				while (true) {		
					if (!rowIterateMatrix.hasNext())
						break;
					rowX = (HSSFRow)rowIterateMatrix.next();
					cellX = rowX.getCell(v);
					if (cellX==null || cellX.getStringCellValue().isBlank())
						continue;
					if (cellX.getStringCellValue().equals(refX.formatAsString())) {
						//set direction flag 
						if (dir.isBlank())
							dir = rowX.getCell(v+1).toString();
						//get row of values into matrix
						while (true) { 
							cellX = rowX.getCell(pv);
							if (cellX==null||cellX.getStringCellValue().isBlank()||cellX.getStringCellValue().equals(END))
								break;
							if (cellX.getCellType()==CellType.NUMERIC)
								if (DateUtil.isCellDateFormatted(outcell))
									yaxis.add(cellX);
								else 
									yaxis.add(cellX.getNumericCellValue());
							else 
								yaxis.add(cellX.getStringCellValue());	
							pv++;
						}
						if (yaxis.size()<1)
							throw new AlsCustomException(annotation+" Cross-Tab has no data at "+refX.formatAsString());
						if (xint==0) { 
							xint = yaxis.size();
							xaxis = yaxis;
							yaxis = new ArrayList<Object>();
							pv = vcol;
							if (xtabcode.length>1) {
								refX = new CellReference(xtabcode[1]);
								rowIterateMatrix = prsheet.rowIterator();  
								rowIterateMatrix.next();//row tracker starts at 1 and shall be running number 
							}
							else 
								break;
						}else if (yaxis.size()>0) {
							yint = yaxis.size();
							break;
						}
					}
				}	
				if (xaxis==null)
					throw new AlsCustomException("CROSS TAB NO SUCH INPUT AT : "+annotation);
				if (yint<2)yint=2; if (xint<2)xint=2;
				canvas = new Object[xint][yint];
				if (yint>xint)xint=yint;if(xint>yint)yint=xint;
				matrix = new Object[xint][yint];
				//feed axis indexes into build matrix 
				for (int x=0;x<xaxis.size();x++) {
					matrix[x][1]=xaxis.get(x);
				}
				if (!yaxis.isEmpty())
				for (int y=0;y<yaxis.size();y++) {
					matrix[y][0]=yaxis.get(y);		
					 
				}
				Cell buffercell =  null;
				//build prValueTray to hold the one or two rows of values above prValues 
				
				for (int p=vcol;p<99;p++) { //going each prValue axis set directly above need not in order
					if (prValue == null)
						break;
					if (prValue.toString().equals(END))
						break;
					if (prValue.toString().equals("NULL")) {
						prValue.setBlank();
					}
					int intRowNo = rowPr.getRowNum();
					int xm=0;  //to hold position within matrix by axis
					int ym = 0;
					int breakflag = 0;
					for (int t=0;t<xtabcode.length;t++) { //reading row above for index values
						 intRowNo--; 
						 HSSFRow bufferrow = prsheet.getRow(intRowNo);;
						 buffercell = bufferrow.getCell(p);
						 if (buffercell==null || buffercell.getStringCellValue().isBlank())
							 break;
						//look up row axis position
						 Object o = null;
						 if (buffercell.getCellType()==CellType.NUMERIC)
								if (DateUtil.isCellDateFormatted(buffercell))
									o=buffercell;
								else 
									o=buffercell.getNumericCellValue();
							else 
								o=buffercell.getStringCellValue();	
						for (int y=0;y<yint;y++) {
							if (o.equals(matrix[y][0])) {
								ym=y;
								breakflag++;
								break;
							}
						}
						if (breakflag==2)
							break;
						for (int x=0;x<xint;x++) {
							if (o.equals(matrix[x][1])) {
								xm=x;
								breakflag++;
								break;
							}
						if (breakflag==2)
								break;
						}
					 } 
					//build values matrix
					if (canvas[xm][ym]==null)
						canvas[xm][ym]= prValue;
					nc++;
					prValue = rowPr.getCell(nc);
				}
				//assign whole matrix to outcell and quit
				//direction if first axis is Y (left vertical) then toggle rotate canvas 
				if (dir.equals(DOWN)) {
					for (int x=0;x<canvas.length;x++) {
						mrow = outsheet.getRow(ref.getRow()+x);
						//if row is null, create it first
						if (mrow==null)
							mrow = outsheet.createRow(ref.getRow()+x);
						Object o;
						for (int y=0;y<canvas[0].length;y++) { 
							o = canvas[x][y];
							if (o==null || o.toString().equals(END))
								continue;
							String str = o.toString(); 
							mcell = mrow.getCell(ref.getCol()+y);
						 	//if cell null, create it first
						 	if(mcell==null)
						 		mcell = mrow.createCell(ref.getCol()+y);
							if (str.equals("NULL")||str.isEmpty()) {
								mcell.setBlank();
								continue;
							}
						 	double d = Double.valueOf(str).doubleValue(); 
							mcell.setCellValue(d);
						}
					}
				} else {
					for (int y=0;y<canvas[0].length;y++) {
						mrow = outsheet.getRow(ref.getRow()+y);
						Object o;
						for (int x=0;x<canvas.length;x++) { 
							o = canvas[x][y];
							if (o==null || o.toString().equals(END))
								continue;
							String str = o.toString(); 
							mcell = mrow.getCell(ref.getCol()+x);
						 	//if cell null, create it first
						 	if(mcell==null)
						 		mcell = mrow.createCell(ref.getCol()+x);
							if (str.equals("NULL")||str.isEmpty()) {
								mcell.setBlank();
								continue;
							}
						 	double d = Double.valueOf(str).doubleValue(); 
							mcell.setCellValue(d);
						}
					}
				}
				
				System.out.println(canvas.length+" CROSS TAB CANVAS: "+annotation);
				break;
			}
			nc++;
			prValue = rowPr.getCell(nc);
			//
		}//LOOP OUTCELL WITHIN OUTROW ++++
	} 
}

/**
 * Shifting to right is done twice. 
 * 1. Move end column to new column
 * 2. Move previous column to end column
 * 3. Hide/Unhide if so for the end-column
 * @param shiftsheet
 * @param startcolumn - 1 relative value - is previous before end. 
 * 						2 is the end column you want to move actually
 * @param endcolumn   - 3 relative value - getLastCellNum() end+1 is new column.
 */
private static void pushColumnsRight(HSSFSheet shiftsheet, int startcolumn, int endcolumn) {
	boolean hidden=false;
	if (shiftsheet.isColumnHidden((startcolumn)+1))
			hidden=true;
	int nocols = endcolumn - startcolumn;
	int a = endcolumn - 1;
	int b = endcolumn;
	for (int n=0; n<nocols; n++) {
		shiftOneColumnRight(shiftsheet, a, b);
		if (a==startcolumn)
			break;
		b--;
		a--;
	}
	if (hidden) {
		shiftsheet.setColumnHidden(startcolumn+1, false);
		shiftsheet.setColumnHidden(endcolumn, true);
	}
}

/**
 * Copy over one column only. Call only from shiftColumnsRight twice
 * @param shiftsheet
 * @param fromcolumn - 2,1 relative
 * @param tocolumn   - 3,2 relative
 */
static void shiftOneColumnRight(HSSFSheet shiftsheet, int fromcolumn, int tocolumn) { 
	if (tocolumn-fromcolumn>1) {
		System.out.println( "CANNOT SHIFT SINGLE COLUMN MORE THAN "+(tocolumn-fromcolumn));
		return;
	}
	int rowno = 0;
	for (int r = 0;r<shiftsheet.getLastRowNum();r++) {
		rowno++; 
		HSSFRow nowrow = shiftsheet.getRow(rowno);
		if (nowrow==null) {
			System.out.println("This row is null : "+rowno);
			continue;
		}
		//from-cell = 22 
		HSSFCell fromcell = nowrow.getCell(fromcolumn);
		if (fromcell==null)
			fromcell = nowrow.createCell(fromcolumn);
		//to-cell = 23
		HSSFCell tocell = nowrow.createCell(tocolumn);
		//copy over from-cell to to-cell
		copyover(shiftsheet, fromcell, tocell);
		if (conditions!=null)
			pushConditionalFormatting(fromcell, RIGHT);
	}			
}

private static void copyover(HSSFSheet shiftsheet, HSSFCell lastcell, HSSFCell newcell) {
	newcell.setCellStyle(lastcell.getCellStyle());
	newcell.setBlank();//remove formula that may remain during copy to.
	if (lastcell.getCellType()==CellType.NUMERIC) {
		newcell.setCellValue(lastcell.getNumericCellValue()); 
	}else if (lastcell.getCellType()==CellType.FORMULA) {
		String shifted = copyFormula(shiftsheet, lastcell.getCellFormula(), 1, 0);
		newcell.setCellFormula(shifted); 
	}else if (lastcell.getCellType()==CellType.STRING) {
		newcell.setCellValue(lastcell.getRichStringCellValue()); 
	}
}

private static HSSFCell pushRowDown(HSSFSheet outsheet, int nc, HSSFRow rowPr, HSSFCell outcell, int downrow, int c) {
	HSSFRow outwriterow;
	HSSFCell peek = rowPr.getCell(nc+1);
	if (peek!=null && !peek.toString().equals(END) && peek.getCellType()!=CellType.BLANK){
		if (outcell.getCellType()==CellType.FORMULA) {
			// INSERT row
			HSSFRow orirow = outsheet.getRow(downrow-1); 
		    int startRow = downrow;
		    int rowNumber = 1;
			int lastRow = outsheet.getLastRowNum(); 
		    outsheet.shiftRows(startRow, lastRow, rowNumber, true, true);
		    outwriterow = outsheet.createRow(startRow); 
		    
		    for (int st=0;st<99;st++){//copy styles
		    	HSSFCell ori = orirow.getCell(st);
		    	HSSFCell niu = outwriterow.createCell(st);
		    	if (ori!=null) {
			    	niu.setCellStyle(ori.getCellStyle()); 
			    	if (ori.getCellType()==CellType.FORMULA) {
			    		String shifted = copyFormula(outsheet, ori.getCellFormula(), 0, 1);
			    		niu.setCellFormula(shifted);
			    	}
		    	}
		    }
		    
		    if (conditions!=null) 
		    	pushConditionalFormatting(outcell, DOWN);
				System.out.println("Conditional Formatting at ROW"); 
		    
		    outcell = outwriterow.getCell(c); 
		}
		
	}
	return outcell;
}

private static String copyFormula(Sheet sheet, String formula, int coldiff, int rowdiff) {
	String old = formula;
	  org.apache.poi.ss.usermodel.Workbook workbook = sheet.getWorkbook();
	  HSSFEvaluationWorkbook evaluationWorkbook = HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook); 

	  Ptg[] ptgs = FormulaParser.parse(formula, (FormulaParsingWorkbook)evaluationWorkbook, 
	   FormulaType.CELL, sheet.getWorkbook().getSheetIndex(sheet));

	  for (int i = 0; i < ptgs.length; i++) {
	   if (ptgs[i] instanceof RefPtgBase) { // base class for cell references
	    RefPtgBase ref = (RefPtgBase) ptgs[i];
	    if (ref.isColRelative())
	     ref.setColumn(ref.getColumn() + coldiff);
	    if (ref.isRowRelative())
	     ref.setRow(ref.getRow() + rowdiff);
	   }
	   else if (ptgs[i] instanceof AreaPtgBase) { // base class for range references
	    AreaPtgBase ref = (AreaPtgBase) ptgs[i];
	    if (ref.isFirstColRelative())
	     ref.setFirstColumn(ref.getFirstColumn() + coldiff);
	    if (ref.isLastColRelative())
	     ref.setLastColumn(ref.getLastColumn() + coldiff);
	    if (ref.isFirstRowRelative())
	     ref.setFirstRow(ref.getFirstRow() + rowdiff);
	    if (ref.isLastRowRelative())
	     ref.setLastRow(ref.getLastRow() + rowdiff);
	   }
	  }
	  formula = FormulaRenderer.toFormulaString((HSSFEvaluationWorkbook)evaluationWorkbook, ptgs);
	  System.out.println(old+" >>formula>> "+formula);
	  return formula;
}

private static void pushConditionalFormatting(HSSFCell cell, String dir) {
    for (int idx = 0; idx < conditions.getNumConditionalFormattings(); idx++) {
    	HSSFConditionalFormatting cf = conditions.getConditionalFormattingAt(idx);
    	List<CellRangeAddress> cra = Arrays.asList(cf.getFormattingRanges());  
	    for (CellRangeAddress c : cra) { 
	        int lastcol = c.getLastColumn(); 
	        int lastrow = c.getLastRow(); 
	        if (dir==RIGHT && cell.getColumnIndex()==lastcol) {
	        	//copy condition to next column cell
	        	c.setLastColumn(lastcol+1);
	        }else if (dir==DOWN && cell.getRowIndex()-2==lastrow) {
	        			//copy condition next row cell
	        			c.setLastRow(lastrow+1);
	        }
	    }
    } 
}

private static void checkConditionalFormatting(){
	HSSFSheetConditionalFormatting results = outputsheet.getSheetConditionalFormatting();
    if (results!=null)
    	conditions=results;
}

/** 
 * CROSSTAB feature - for '#?..;' advice in WHERE of Process Sheet 
 * When it is flagged during MakeFullSQL,
 * it shall store the SQL in crosstabMATRIX, then minus the SELECTion 
 * SELECT - name in the row
 * WHERE  - add conditions to address set
 * replace result with progressive values 
 * Paint not done here
 * 
 * EXAMPLE: //a.DateTrx = [C8] AND (b.Value = [N4] OR b.Value = [O4])
 * 1. WHERE .. AND b.Value=(content of Cell N4) AND a.DateTrx<(content in Cell C8)
 * 2. extra WHERE SQL add is fully qualified
 * 3. NOTE Painting is thus same taking from Process Values.
 * 4. Later we can do 'jumping' cells as stipulated by address gap.
 * 
 * This means the Values in Process Sheet are written twice. 
 * Second pass is the right one.
 * Programmer sanity first :D
 *
 * @param processsheet
 * @param outsheet
 * @throws SQLException 
 */ 
private static void crosstab(HSSFSheet processsheet, HSSFSheet outsheet) throws AlsCustomException, SQLException {
	/**
	 *  #?a.DateTrx = [C8] AND (b.Value = [N4] OR b.Value = [O4]);
	 *  PARSE TO
	 *  AND a.DateTrx < '22/02/22' AND (b.Value = '601000' OR b.Value = '602000')
	 *  
	 * TODO
	 *  **/
	//loop between last row write and last row read i.e. 1 to 23
	//to look for cross tab statements in WHERE 
	String ori;  
	StringBuilder activeSQL; 
	HSSFRow rowPr;
	HSSFCell cellPr;
	String SELECT;    
	String cell5value; 
	StringBuilder hash2comma;
	HSSFCell selectcell; 
	List<List<Object>> result; 
	
	/************************* TAKE SQL WITHOUT SELECTION, START FROM ***********/
	ori = crossTabHoldSQL.trim(); 
	if (ori.indexOf(" FROM")==-1)
		throw new AlsCustomException("NO ' FROM' IN ORIGINAL SQL "+ori);

	/*********** LOOK FOR EACH CROSSTAB THRU EACH ROW WITHIN LAST SET ************/
	for (int a=lastrowwrite;a<lastrowread;a++) {
		activeSQL = new StringBuilder(ori.substring(ori.indexOf(" FROM")));
		
		rowPr = processsheet.getRow(a);
		cellPr = rowPr.getCell(5);
		if (cellPr==null)
			continue;
		if (cellPr.getStringCellValue().contains("#?")) {

			/*********************** REMOVE MARKINGS #? ... ; ********************/
			cell5value = cellPr.getStringCellValue() ;
			hash2comma = new StringBuilder(cell5value.substring(2,cell5value.length()-1));
			hash2comma.toString().trim(); 
			
			/************************* SELECT NAME ************************/
			selectcell = rowPr.getCell(2);// third column
			if (selectcell==null)
				throw new AlsCustomException("NO NAME IN SELECT ROW OF PROCESS SHEET");
			SELECT = selectcell.getStringCellValue(); 
			
			/**************** ACTIVE SQL BEGIN ***************************/
			activeSQL = new StringBuilder("SELECT "+SELECT).append(activeSQL);
			if (ori.contains("WHERE"))
				activeSQL.append(" AND ");
			
			/************* PARSE BRACKET TAGS REFERING OUTPUT SHEET *********/
			StringBuilder h2cbuffer = parseBracketTag(outsheet, hash2comma);
			
			activeSQL.append(h2cbuffer);
			crosstab=false; int v=vcol;
			result = executeSQL(activeSQL) ;
			if (result==null) { 
				continue;
			}
			for (List<Object>ctresult:result) {

				for (Object retValue:ctresult) {  
					//set values to Process Row second time
					 
						if (retValue==null)
							retValue="NULL";
						if (rowPr.getCell(v)==null)
							rowPr.createCell(v);
						 	cellPr = rowPr.getCell(v);  	
						String cellValue = dataFormatter.formatCellValue(cellPr); 
						if (cellValue.equals(END) && v>vcol) //avoid accident with bottom END on first 7th column
							return;
						//check if numeric
						 if (retValue instanceof String) { //STRING TYPE ++++
							 if (String.valueOf(retValue).equals(cellPr.toString()))
								continue;//same string, stop writing further
								else {	 	
									cellPr.setCellValue((String)retValue);
									System.out.println(cellPr.getAddress()+" CROSSTAB String "+retValue);
								}
						 }else if (retValue instanceof Timestamp) {
							 if (retValue.equals(cellPr.getLocalDateTimeCellValue()))
									continue;//same string, stop writing further
									else {	 	
										cellPr.setCellValue((Timestamp)retValue); 
										System.out.println(cellPr.getAddress()+" CROSSTAB Timestamp "+retValue);
									}
						 }
						 else {//NUMERIC ++++
							//double celltotal = cellPr.getNumericCellValue();
						 	String str = retValue.toString(); 
						 	double d = Double.valueOf(str).doubleValue(); 
							cellPr.setCellValue(d); //for both cases
							System.out.println(cellPr.getAddress()+" CROSSTAB Numeric "+retValue);		
						 }
						 v++;
				} 
			}/***************** FINAL IN CROSSTAB MARKED #? .. ; FOR LOOP ************************/
		}
	} 		
	crossTabHoldSQL = "";
	crosstab=false;
}


/*
 * Set DB returned RESULTS from Full SQL down wise by row
 * Down wise because the RESULTS matrix orientation
 */
private static void set(List<List<Object>> results, HSSFSheet prsheet) {
	int v = vcol;
	for (List<Object> returning:results) { 
		Iterator<Row> rowIteratePr = prsheet.rowIterator();   

		for (int r=0;r<lastrowwrite;r++) {
			rowIteratePr.next();//row tracker starts at 1 and shall be running number
		}
		boolean done = setReturnValues(prsheet,rowIteratePr, returning, v); 
		if (done)
			break;
		v++; //column to write returned values starting from 7 to the END
	}
}

/*	SET return values onto Process Sheet Col V++ till hit bottom End
 * If done return true. False means continue looping by parent
 */
private static boolean setReturnValues(HSSFSheet prsheet, Iterator<Row> rowIteratePr, List<Object> returning, int v) { 
	for (Object retValue:returning) {
		if (retValue==null)
			retValue="NULL";
		if (!rowIteratePr.hasNext())
			return false;
		HSSFRow rowPr = (HSSFRow)rowIteratePr.next();
		if (rowPr.getCell(v)==null)
			rowPr.createCell(v);
		HSSFCell cellPr = rowPr.getCell(v);  	
		
		//by pass Cross Tab set
		HSSFCell checkXT = rowPr.getCell(5);
		if (checkXT !=null && checkXT.getStringCellValue().contains("#?"))
			continue;
		
		String cellValue = dataFormatter.formatCellValue(cellPr); 
		if (cellValue.equals(END) && v>vcol) //avoid accident with bottom END on first 7th column
			return true;
		
		//check if numeric
		 if (retValue instanceof String) { //STRING TYPE ++++
			 if (String.valueOf(retValue).equals(cellPr.toString()))
				continue;//same string, stop writing further
				else {	 	
					cellPr.setCellValue((String)retValue);
					System.out.println(cellPr.getAddress()+" Address is String "+retValue);
				}
		 }else if (retValue instanceof Timestamp) {
			 if (retValue.equals(cellPr.getLocalDateTimeCellValue()))
					continue;//same string, stop writing further
					else {	 	
						cellPr.setCellValue((Timestamp)retValue); 
						System.out.println(cellPr.getAddress()+" Address is Timestamp "+retValue);
					}
		 }
		 else {//NUMERIC ++++
			//double celltotal = cellPr.getNumericCellValue();
		 	String str = retValue.toString(); 
		 	double d = Double.valueOf(str).doubleValue(); 
			cellPr.setCellValue(d); //for both cases
			System.out.println(cellPr.getAddress()+" Address is Numeric "+retValue);		
		 } 
			
			//copy over output Style
			if (v>vcol) {
				HSSFCell b4cell = rowPr.getCell(vcol);
				cellPr.setCellStyle(b4cell.getCellStyle());
			}
	 }
	return false;
}

private static List<List<Object>> executeSQL(StringBuilder SQL) throws AlsCustomException, SQLException {
	List<List<Object>> result =  new ArrayList<List<Object>>();
	try {
		 Connection conn = DriverManager.getConnection(DB_URL, USER, PASS);
		 Statement stmt = conn.createStatement();
         ResultSet rs = stmt.executeQuery(SQL.toString());
         if (!rs.isBeforeFirst() ) {  
			System.out.println(SQL.toString()+" RETURNS NOTHING");
			return null;
		}
         ResultSetMetaData rsmd = rs.getMetaData();
         while (rs.next()) {
    			List<Object> retValue = new ArrayList<Object>();
    			for (int i=1; i<=rsmd.getColumnCount(); i++) {
    				Object obj = rs.getObject(i);
        			if (rs.wasNull())
        				retValue.add(null);
        			else
        				retValue.add(obj);
    			}
    			result.add(retValue);
    		}
		conn.close();
	}
	 catch(Exception e) {
		String analysis = formatException(e,SQL.toString());
		throw new AlsCustomException(analysis);
	} 
	return result;
}

private static String formatException(Exception e, String sqlstring) {
	String exception = e.toString(); 
	String error = "";
	String error2 = "";
	int from = sqlstring.indexOf(" FROM ");
	int join = sqlstring.indexOf(" JOIN ");
	int where = sqlstring.indexOf(" WHERE ");
	int pos1 = exception.indexOf("\"");
	int pos2 = exception.indexOf("\"", exception.indexOf("\"") + 1);
	if (pos1+pos2>0)
		error = exception.substring(pos1+1,pos2); 
	
	//Second ERROR marking by position
	pos1 = exception.indexOf("Position:")+10;
	pos2 = exception.length();

	String test = exception.substring(pos1,pos2).trim();
	if (test.matches("[0-9]+")) {  
		Integer sub = Integer.parseInt(test)-1;
		Integer sub2 = sub+8;
		if (sub2 > sqlstring.length())
			sub2 = sqlstring.length(); 
		error2 = sqlstring.substring(sub,sub2);		
	}		
	 
		//
	String selectpart = "<br><font color=\"grey\">" + sqlstring.substring(0,from);
	String frompart = "<font color=\"blue\">" + sqlstring.substring(from,join);
	String joinpart = "<font color=\"505050\">" + sqlstring.substring(join,where);
	String wherepart = "<font color=\"grey\">" + sqlstring.substring(where); 
	//
	sqlstring = exception+"<p><br>SQL :"+sqlcount+"<br>"+selectpart+frompart+joinpart+wherepart;
	if (error.length()>error2.length())
		sqlstring = sqlstring.replace(error,"<mark>"+error+"</mark>"); 
	else
		sqlstring = sqlstring.replace(error2,"<mark>"+error2+"</mark>"); 
	//
	return sqlstring;
}
/**
 * Parse Bracket Tag refers to Output Sheet values by Address
 * Can be used as CrosstTab, inject for sub selection or in complete SQL
 * @param outsheet
 * @param hash2comma
 * @return
 */
private static StringBuilder parseBracketTag(HSSFSheet outsheet, StringBuilder hash2comma) throws AlsCustomException {
	String C8;
	String h2cbuffer = hash2comma.toString(); 
	//parse main statement
	//replace address [C8] with referred displayed value 	
	/*************** CHECK FOR (C8) EXISTENCE ****************************/
	int open = h2cbuffer.indexOf("[");
	int close = h2cbuffer.indexOf("]");
	HSSFCell outcell;
	while (open+close>0) {						
		/************* C8 ... N4  ******************************/
		C8 = h2cbuffer.substring(open+1, close); //C8
		h2cbuffer = h2cbuffer.replaceFirst("\\[","'").replaceFirst("\\]","'");//cutting off when taken from front
		
		//get  CROSSTAB address in Output Sheet
		CellReference ref = new CellReference(C8);
		//get value from address put into SQLRow outrow = outsheet.getRow(ref.getRow());
		Object valueC8 = null;
		HSSFRow outrow = outsheet.getRow(ref.getRow());
		if (outrow==null)
			throw new AlsCustomException("CROSSTAB - OUTPUT SHEET NULL ROW: [TAG] is "+C8);
		   	outcell = outrow.getCell(ref.getCol());
		if (outcell == null) {
			String exception  = hash2comma.toString().replace(C8,"<b><mark>"+C8+"</mark></b>"); 
			throw new AlsCustomException("CROSSTAB - OUTPUT SHEET [TAG] IS NULL: <p><br><font color=\"grey\">"+exception);
		}
			outcell = outrow.getCell(ref.getCol());  
		if (outcell.getCellType()==CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(outcell)) {
				valueC8 = outcell; 
	            h2cbuffer = h2cbuffer.replaceFirst(C8,valueC8.toString());
	        } else {
	        	DataFormatter formatter = new DataFormatter(); 
	        	h2cbuffer = h2cbuffer.replaceFirst("'"+C8+"'",formatter.formatCellValue(outcell));
	        }
		}else  
			h2cbuffer = h2cbuffer.replaceFirst(C8,outcell.getStringCellValue().trim());	
		
		/*************** CHECK FOR NEXT (C8) EXISTENCE ****************************/
		open = h2cbuffer.indexOf("[");
		close = h2cbuffer.indexOf("]");
	} /************************ END METHOD *******************/
	return new StringBuilder(h2cbuffer);
}

/**
 * Formulate a Fully Qualified SQL Statement for Execution
 * @param prsheet
 * @return
 */
private static StringBuilder makeFullSQLStatement(HSSFSheet prsheet) {		
	StringBuilder fullSQL 		= new StringBuilder(); 
	StringBuilder SELECT 		= new StringBuilder();
	StringBuilder TABLE 		= new StringBuilder();
	StringBuilder JOIN 			= new StringBuilder(); 
	StringBuilder WHERE 		= new StringBuilder();
	
	boolean gothruall 			= false; 
	Iterator<Row> rowIteratePr 	= prsheet.rowIterator(); 
	
	for (int r=0;r<lastrowwrite;r++) {
		rowIteratePr.next();//row tracker starts at 1 and shall be running number
	} 
	HSSFRow rowPr = null;
	while (rowIteratePr.hasNext()) { //loop thru every row
		rowPr = (HSSFRow)rowIteratePr.next();	
		if (gothruall) { // #########################################################################
			// ********** CHECK IF NEW SQL NEXT, THEN RETURN TO WRITE FIRST AND COME BACK *****
			if ((rowPr.getCell(3)!=null) && rowPr.getCell(3).toString().endsWith(" a")) { 	//meets another new set of Tables
				gothruall=false; //implied  		//..so exit to process SQL then return 
				
				// ############ FORM FULL SQL AND REPLACE #1,#2,.. WITH PARAMETERS					 
				fullSQL = selectJoinWhereSQL(fullSQL, SELECT, TABLE, JOIN, WHERE); 
				lastrowread = rowPr.getRowNum();
				return fullSQL;						
			}//*********************************************************************************
			if (rowPr.getCell(2)!=null && !rowPr.getCell(2).toString().isBlank())
				SELECT.append(","+rowPr.getCell(2).toString());
			if (rowPr.getCell(3)!=null) {
				if (rowPr.getCell(3).toString().contains("JOIN")) // INNER/LEFT/RIGHT/OUTER JOINS
					TABLE.append(" "+rowPr.getCell(3).toString());
				else {
					if (!rowPr.getCell(3).toString().isEmpty())
						TABLE.append(","+rowPr.getCell(3).toString());						
				}
			}
			if (rowPr.getCell(4)!=null) {
				if (!rowPr.getCell(4).toString().isEmpty())
					JOIN.append(" "+rowPr.getCell(4).toString());	
			}
			if (rowPr.getCell(5)!=null) {
				if (!rowPr.getCell(5).toString().isEmpty())
					if (rowPr.getCell(5).toString().contains("#?")) 
						//CROSSTAB feature - parse add snip SQL and pass more during Results
						crosstab = true;
					WHERE.append(" "+rowPr.getCell(5).toString());	
			}				
			continue; //picking up more SQL bits
		} else // ###################################################################################
			SELECT = new StringBuilder(rowPr.getCell(2).toString());
			
		if (rowPr.getCell(3)!=null)
			TABLE = new StringBuilder (rowPr.getCell(3).toString());
		rowPr.getRowNum();
		if (rowPr.getCell(4)!=null)
			JOIN = new StringBuilder (rowPr.getCell(4).toString());
		 
		if (rowPr.getCell(5)!=null) {
			if (rowPr.getCell(5).toString().contains("#?")) 
				//CROSSTAB feature - parse add snip SQL and pass more during Results
				crosstab = true;
			WHERE = new StringBuilder (rowPr.getCell(5).toString());
		}
		
		if(SELECT != null && TABLE != null && JOIN != null&& WHERE != null) 
		{
			// ############ FORM FULL SQL AND REPLACE #1,#2,.. WITH PARAMETERS					 
			fullSQL = selectJoinWhereSQL(fullSQL, SELECT, TABLE, JOIN, WHERE); 
			lastrowread = rowPr.getRowNum() + 1;
			if (!rowIteratePr.hasNext())
				gotSQL=false; //you run out of SQL lines in Process Sheet, so don't come back ! :)
			return fullSQL;
		}
		
		if (!gothruall)
			if (TABLE.toString().endsWith(" a")) {
				gothruall=true;
			}
	}
	lastrowread = rowPr.getRowNum();
	// ############ FORM FULL SQL AND REPLACE #1,#2,.. WITH PARAMETERS
	fullSQL = selectJoinWhereSQL(fullSQL, SELECT, TABLE, JOIN, WHERE);
	String complete = crossTabRemoveHash(fullSQL);
	if (!rowIteratePr.hasNext())
		gotSQL=false; //you run out of SQL lines in Process Sheet, so don't come back ! :)
	return new StringBuilder(complete);
}

static String crossTabHoldSQL = "";
private static String crossTabRemoveHash(StringBuilder fullSQL) {
	String complete  = fullSQL.toString();
	if (crosstab) { //CROSSTAB feature
		//repeat SQL with column progression and add to results  
		complete =complete .replaceAll("#\\?.*?;", "");
		crossTabHoldSQL=complete;  
	}
	return complete;
}

private static StringBuilder selectJoinWhereSQL(StringBuilder fullSQL, StringBuilder SELECT, StringBuilder TABLE,
		StringBuilder JOIN, StringBuilder WHERE) {
	fullSQL.append("SELECT "+SELECT).append(" FROM "+TABLE)
	.append(" WHERE "+WHERE)
	.append(((JOIN.toString().isBlank())?"":" AND ")).append(JOIN);
	fullSQL = replaceParameterTag(fullSQL, from, to);
	return fullSQL;
}

private static StringBuilder replaceParameterTag(StringBuilder whereCondition, String from, String to) {
	String s = "";
	if (whereCondition.toString().contains("#")){
		s = whereCondition.toString().replace("#1","\'"+from+"\'"); 
		whereCondition = new StringBuilder(s.replace("#2","\'"+to+"\'"));
		s = whereCondition.toString();
		whereCondition = new StringBuilder(s.replace("#3","\'"+AccountGroup+"\'"));
	}
	return whereCondition;
} 

private static void setAddress(HSSFSheet insheet,HSSFSheet prsheet, HSSFSheet outsheet) {		
	Iterator<Row> rowIteratePr = prsheet.rowIterator(); 
	boolean nothing = true;
	rowIteratePr.next();//skip label header
	while (rowIteratePr.hasNext()) { //loop thru every row
		HSSFRow rowPr = (HSSFRow)rowIteratePr.next();	
		HSSFCell check = rowPr.getCell(1);
		if (check==null)
			break;// EOF - end of ROW() Column
		HSSFCell annotate = rowPr.getCell(0); 
		if (annotate==null || annotate.toString().isBlank())
			continue;
		String annotateValue = annotate.getRichStringCellValue().toString();

		HSSFCell addressPr = rowPr.getCell(6);//Cell to store address of input annotation
		if (addressPr==null) {
			rowPr.createCell(6);  			//initialize cell if null
			addressPr=rowPr.getCell(6);
		}
		//Search address in input sheet
		Iterator<Row> rowIterateIn = insheet.rowIterator(); //go thru input rows
		//
		while(rowIterateIn.hasNext()) {  
			HSSFRow rowIn = (HSSFRow)rowIterateIn.next();	
			Iterator<Cell> cellIterateIn = rowIn.cellIterator();
			while(cellIterateIn.hasNext()) {
				HSSFCell cellIn = (HSSFCell) cellIterateIn.next();
				if (cellIn.getCellType()==CellType.NUMERIC)
					continue;
				String cellInValue = cellIn.getRichStringCellValue().toString();
				
				if (cellInValue.equals(annotateValue)) {
					nothing = false; 
					
					//get output style
					CellAddress address = cellIn.getAddress();
					HSSFRow outrow = outsheet.getRow(address.getRow());
					if (outrow==null)
						outrow = outsheet.createRow(address.getRow());
					HSSFCell outcell = outrow.getCell(address.getColumn()); 
					if (outcell==null)
						outcell = outrow.createCell(address.getColumn());
					HSSFCell value = rowPr.getCell(8);
					if (value==null) {
						rowPr.createCell(8);
						value = rowPr.getCell(8);
					}
					value.setCellStyle(outcell.getCellStyle());
					
					addressPr.setCellStyle(outcell.getCellStyle());
					
					//set to process sheet
					addressPr.setCellValue(cellIn.getAddress().toString()); 
					//  						 
					break;
				}	
			}
		}
	}
	if (nothing)
		System.out.println("Nothing was painted. Do INPUT sheet again");
	
}

private static boolean putImage(HSSFSheet input, HSSFSheet output) throws IOException {  
	//look for @IMAGE=AD_ORG.VALUE.JPEG 
	boolean gotlogo = false;
	final String imagestart = "@IMAGE=";
	String image = "";
	CellAddress location = null;
	//Search address in input sheet
	Iterator<Row> rowIterateIn = input.rowIterator(); //go thru input rows 
	while(rowIterateIn.hasNext()) {  
		if (gotlogo)
			break;
		HSSFRow rowIn = (HSSFRow)rowIterateIn.next();	
		Iterator<Cell> cellIterateIn = rowIn.cellIterator();
		while(cellIterateIn.hasNext()) {
			HSSFCell cellIn = (HSSFCell) cellIterateIn.next();
			if (cellIn.getCellType()==CellType.NUMERIC)
				continue;
			String cellInValue = cellIn.getRichStringCellValue().toString();
			
			if (cellInValue.startsWith(imagestart)) {
				//get reference address
				location = cellIn.getAddress(); 
				gotlogo = true;
				//get image string
				image = cellInValue;
				break;
			}
		}
	}
	if (gotlogo) {
		String[] splitimagestring = image.split("\\.");
		  String Org = "MOES";
		  final String imagepath = path+ "/"+Org+"."+splitimagestring[2];	
		  File gotF = new File(imagepath);
	  	  if(!gotF.exists())
	  		  return false;
		  final FileInputStream stream =  new FileInputStream(imagepath);
		  
          final CreationHelper helper = workbook.getCreationHelper();
          final Drawing<?> drawing = output.createDrawingPatriarch();

          final ClientAnchor anchor = helper.createClientAnchor();
          anchor.setAnchorType( ClientAnchor.AnchorType.MOVE_AND_RESIZE);
          String pictype = splitimagestring[2].toUpperCase();
          int PIC_TYPE = 0;
          if (pictype.equals("JPEG")||pictype.equals("JPG"))
        	  PIC_TYPE = HSSFWorkbook.PICTURE_TYPE_JPEG;
          else if (pictype.equals("PNG"))
        	  PIC_TYPE = HSSFWorkbook.PICTURE_TYPE_PNG;
          else if (pictype.equals("PICT"))
        	  PIC_TYPE = HSSFWorkbook.PICTURE_TYPE_PICT;
          final int pictureIndex =
          workbook.addPicture(IOUtils.toByteArray(stream), PIC_TYPE);
 
          int row = location.getRow();
          int col = location.getColumn();
          anchor.setCol1(col);
          anchor.setRow1(row);
          anchor.setRow2(row+5);
          anchor.setCol2(col);
          final Picture pict = drawing.createPicture( anchor, pictureIndex );
          pict.resize(1.0,1.0);
	}			
      return gotlogo;
	
}

/**
 * Address Column and Values Column in Process Sheet
 * @param prsheet
 * @throws IOException
 */
private static void clean(FileInputStream file)throws IOException { 
	processsheet = workbook.getSheet("Process"); 
	outputsheet = workbook.getSheet("Output");

	//clone complete sheet from BACKUP Sheet(4)
	int out = workbook.getSheetIndex(outputsheet);
	workbook.setSheetName(out, "Test"); 
	int bak = workbook.getSheetIndex("BACKUP");
	workbook.cloneSheet(bak);
	workbook.setSheetName(bak, "Output");  
	int two = workbook.getSheetIndex("BACKUP (2)");
	workbook.setSheetName(two, "BACKUP");
	workbook.removeSheetAt(out); 

	Iterator<Row> rowIteratePrClear = processsheet.rowIterator(); 		
	rowIteratePrClear.next();//skip label header
	while (rowIteratePrClear.hasNext()) {   

		int col = vcol;
		HSSFRow rowPr  = (HSSFRow)rowIteratePrClear.next();//row 2 until END
		//
		// clear column G of address
		HSSFCell cellAdd = rowPr.getCell(6); 
		
		if (cellAdd!=null) 
			cellAdd.setBlank();
		
		//clear values row
		HSSFCell cellPr = null; 
		while (true) {	
			cellPr = rowPr.getCell(col);
			col++;
			if (col>99)
				break;
			if (cellPr==null)
				break;
			if (cellPr.getCellType()==CellType.NUMERIC)
				cellPr.setBlank();
			if (cellPr.getStringCellValue().equals(END)) 
				break;		
			if (cellPr.getStringCellValue().isBlank())
				continue;
			cellPr.setBlank();	 
		}
	}
	workbookWrite(); 
	
}

private static void workbookWrite() throws IOException {
	FileOutputStream out = new FileOutputStream(File_Directory);
	if(out!=null)
	{
		workbook.write(out);
		out.close();
	}
}



}


class AlsCustomException extends Exception
{
    public AlsCustomException(String message)
    {
        super(message);
    }
}