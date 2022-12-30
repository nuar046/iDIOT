package my.org;

import java.sql.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import com.itextpdf.licensing.base.LicenseKey;
import com.itextpdf.licensing.base.exceptions.LicenseKeyException;
import com.itextpdf.pdfoffice.OfficeConverter;

import java.io.File;
import java.io.IOException;

public class idiot {
	   static final String DB_URL = "jdbc:postgresql://localhost:5432/idempiere";
	   static final String USER = "adempiere";
	   static final String PASS = "adempiere";
	   static final String QUERY = "SELECT * FROM adempiere.ad_menu";
	   final static int vcol = 8;
	   static int sqlcount = 0;
	   static String DOWN = "V";
	   static String LEFT = "<";
	   static String RIGHT = ">";
	   static String END = "END";
	   static String File_Directory = "C:\\Users\\60133\\Downloads\\MOESReport.xls"; 
	   static HSSFWorkbook workbook = new HSSFWorkbook(); 
	   Timestamp DateFrom = null;
	   Timestamp DateTo = null;
	   static HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
		String direction = "";
		static int lastrowwrite = 1;
		static int lastrowread = 0; 
		static String documentno = "";
		static long start = System.currentTimeMillis(); 
		//SET ParameterTag values for #1 From, #2 To	
		static String from = "";
		static String to = "";
		static boolean gotSQL = true;
		
public static void main(String[] args) throws AlsCustomException, SQLException, InterruptedException {
	      // Open a connection
/*
	      try(Connection conn = DriverManager.getConnection(DB_URL, USER, PASS);
	         Statement stmt = conn.createStatement();
	         ResultSet rs = stmt.executeQuery(QUERY);) {
	         // Extract data from result set
	         while (rs.next()) {
	            // Retrieve by column name
	            System.out.print("ID: " + rs.getInt("ad_menu_id"));
	         }
	      } catch (SQLException e) {
	         e.printStackTrace();
	      } 
*/

		FileInputStream file;
		try {
		   System.out.println("Argument count: " + args.length);
		   for (int i = 0; i < args.length; i++) {
		      System.out.println("Argument " + i + ": " + args[i]);
		   }   
			
			file = new FileInputStream(File_Directory);
			workbook = new HSSFWorkbook(file);
			HSSFSheet configsheet = workbook.getSheet("Config");
			HSSFSheet inputsheet = workbook.getSheet("Input");
			HSSFSheet processsheet = workbook.getSheet("Process"); 
			HSSFSheet outputsheet = workbook.getSheet("Output");
			
			workbook.setForceFormulaRecalculation(true); 
			int norecs = 0;
			List<List<Object>> results = null;
			// 
			//clearContent(processsheet);
			//
			//setAddress(inputsheet,processsheet,outputsheet); 
			/*
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
					paint(processsheet,outputsheet);
				}
				 lastrowwrite = lastrowread;
			}
			*/ 
			System.out.println("END");
			//workbookWrite();
			//workbookClose();
			doPDF();
			//
			long finish = System.currentTimeMillis();
			long timeElapsed = finish - start; 
			System.out.print(norecs+" Records. Time Elapsed: "+BigDecimal.valueOf(timeElapsed));
		} 
		catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
 
	/**
	 * Address Column and Values Column in Process Sheet
	 * @param prsheet
	 * @throws IOException
	 */
	private static void clearContent(HSSFSheet prsheet) throws IOException { 
		Iterator<Row> rowIteratePrClear = prsheet.rowIterator(); 
	
		rowIteratePrClear.next();//skip label header
		while (rowIteratePrClear.hasNext()) {  
			int col = vcol;
			HSSFRow rowPr = (HSSFRow)rowIteratePrClear.next();//row 2 until END
			//
			// clear column G of address
			HSSFCell cell1 = rowPr.getCell(6);
			if (cell1!=null)
				cell1.setBlank();
			
			//clear values row
			HSSFCell cellPr = null;
			while (true) {	
				cellPr = rowPr.getCell(col);
				col++;
				if (col>100)
					break;
				if (cellPr==null)
					continue;
				if (cellPr.getCellType()==CellType.NUMERIC)
					cellPr.setBlank();
				if (cellPr.getStringCellValue().equals(END)) 
					break;					
				cellPr.setBlank();
			}
		}
		workbookWrite();
	}
	
	/**
	 * Get Location on Input Sheet as Address Values in Process Sheet Column 2
	 * @param recordIDs 
	 * @param sheet
	 */
	private static void setAddress(HSSFSheet insheet,HSSFSheet prsheet, HSSFSheet outsheet) {			
		Iterator<Row> rowIteratePr = prsheet.rowIterator(); 
		boolean nothing = true;
		rowIteratePr.next();//skip label header
		while (rowIteratePr.hasNext()) { //loop thru every row
			HSSFRow rowPr = (HSSFRow)rowIteratePr.next();	
			HSSFCell check = rowPr.getCell(1);
			if (check==null)
				break;// EOF - end of Annotate Column
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
					String cellInValue = cellIn.getRichStringCellValue().toString();
					
					if (cellInValue.equals(annotateValue)) {
						nothing = false; 
						
						//get output style
						CellAddress address = cellIn.getAddress();
						HSSFRow outrow = outsheet.getRow(address.getRow());
						HSSFCell outcell = outrow.getCell(address.getColumn()); 
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
	
	private static void workbookWrite() throws IOException {
		 
		FileOutputStream out = new FileOutputStream(File_Directory);
		if(out!=null)
		{
			workbook.write(out);
			out.close();
		}
	}
	
	static void workbookClose(HSSFWorkbook wb) throws IOException { 
		workbook.close();
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

	private static StringBuilder selectJoinWhereSQL(StringBuilder fullSQL, StringBuilder SELECT, StringBuilder TABLE,
			StringBuilder JOIN, StringBuilder WHERE) {
		fullSQL.append("SELECT "+SELECT).append(" FROM "+TABLE)
		.append(" WHERE "+WHERE)
		.append(((JOIN.toString().isBlank())?"":" AND ")).append(JOIN);
		fullSQL = replaceParameterTag(fullSQL, from, to);
		return fullSQL;
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
	
	static boolean crosstab;   
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
	 * @throws AlsCustomException 
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
		String C8;
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
	
	/**
	 * Parse Bracket Tag refers to Output Sheet values by Address
	 * Can be used as CrosstTab, inject for sub selection or in complete SQL
	 * @param outsheet
	 * @param hash2comma
	 * @return
	 * @throws AlsCustomException 
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
				h2cbuffer = h2cbuffer.replaceFirst(C8,outcell.getStringCellValue());	
			
			/*************** CHECK FOR NEXT (C8) EXISTENCE ****************************/
			open = h2cbuffer.indexOf("[");
			close = h2cbuffer.indexOf("]");
		} /************************ END METHOD *******************/
		return new StringBuilder(h2cbuffer);
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

	private static void paint(HSSFSheet prsheet, HSSFSheet outsheet) {
		//take column "Address" as ready array
		Iterator<Row> rowIteratePr = prsheet.rowIterator();  
		for (int r=0;r<lastrowwrite;r++) {
			rowIteratePr.next();//row tracker starts at 1 and shall be running number
		}
		while (rowIteratePr.hasNext()) {  
			int nc = vcol;//column VALUEs start at 7
			HSSFRow rowPr = (HSSFRow)rowIteratePr.next(); 
			//search address in insheet
			HSSFCell addressPr = rowPr.getCell(6); 
			if (rowPr.getCell(0)==null)
				break;
			String annotation = rowPr.getCell(0).getStringCellValue()+rowPr.getCell(7).getStringCellValue(); 
			if (addressPr==null|| addressPr.toString().isBlank())
				continue;
			//fetch values from prSheet
			CellReference ref = new CellReference(addressPr.getStringCellValue());
			Row outrow = outsheet.getRow(ref.getRow());
			if (outrow==null)
				outrow = outsheet.createRow(ref.getRow());
			   Cell outcell = outrow.getCell(ref.getCol());
			if (outcell == null) 
				outcell = outrow.createCell(ref.getCol());  
			//
			//holder for row to go DOWN
			int downrow = outrow.getRowNum();
			Row outwriterow = outrow;
			
			//fetch/write VALUEs to outputsheet    
			boolean prValueNotEnd = true;
			HSSFCell prValue = rowPr.getCell(nc);
			while (prValueNotEnd) {  
				if (prValue == null)
					break;
				if (prValue.toString().equals(END))
					break;
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
					outcell = outwriterow.getCell(outcell.getColumnIndex());
					//
					if (outwriterow.getCell(outcell.getColumnIndex())==null)
						outcell = outwriterow.createCell(outcell.getColumnIndex());
					else
						outcell = outwriterow.getCell(outcell.getColumnIndex());
					
				} else if (annotation.endsWith(RIGHT)){ //continue RIGHT-WISE
					if (outrow.getCell(outcell.getColumnIndex()+1)==null)
						outcell = outrow.createCell(outcell.getColumnIndex()+1);
					else
						outcell = outrow.getCell(outcell.getColumnIndex()+1);
				}
				nc++;
				prValue = rowPr.getCell(nc);
				//
			}//LOOP OUTCELL WITHIN OUTROW ++++
		} 
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
	
	private static StringBuilder replaceParameterTag(StringBuilder whereCondition, String from, String to) {
		String s = "";
		if (whereCondition.toString().contains("#")){
			s = whereCondition.toString().replace("#1","\'"+from+"\'"); 
			whereCondition = new StringBuilder(s.replace("#2","\'"+to+"\'"));
		}
		return whereCondition;
	} 
	 
	static String File_Directory_PDF = "";
	/*
	 * 
	 * @throws IOException
	 */
	static void doPDF() throws IOException,LicenseKeyException{ 
	         
        FileInputStream fis = new FileInputStream(File_Directory);
        Workbook wbk = WorkbookFactory.create(fis);
        wbk.removeSheetAt(0); //remove input sheet
        wbk.removeSheetAt(0); //remove process sheet
        wbk.removeSheetAt(0); //remove Reference sheet
        wbk.removeSheetAt(1); //remove Backup sheet
        Sheet sheet = wbk.getSheetAt(0);
		PrintSetup print = sheet.getPrintSetup();
		print.setScale((short) 55);
		print.setPaperSize(PrintSetup.TABLOID_PAPERSIZE);	
		print.setLandscape(true);
        fis.close();
        wbk.write(new FileOutputStream("output.xls")); 
        wbk.close();
        
        String File_Directory_PDF = File_Directory.replace("xls", "pdf");	
		LicenseKey.loadLicenseFile(new File("license.json"));
		OfficeConverter.convertOfficeSpreadsheetToPdf(new FileInputStream("output.xls"), new FileOutputStream(File_Directory_PDF));			
	}
	
	
}


class AlsCustomException extends Exception
{
    public AlsCustomException(String message)
    {
        super(message);
    }
}