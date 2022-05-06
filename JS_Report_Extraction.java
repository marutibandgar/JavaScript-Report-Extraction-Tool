
import java.awt.RenderingHints.Key;
import java.io.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.*;
import java.nio.file.StandardCopyOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
//import java.util.HashMapQuoteCapt;
import java.util.Map.Entry;

import javax.xml.bind.ParseConversionEvent;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFontFormatting;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.*;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.apache.commons.collections4.MultiMap;
import org.apache.commons.collections4.map.MultiValueMap;
import org.apache.commons.compress.compressors.FileNameUtil;
import org.apache.commons.math3.optim.nonlinear.scalar.GoalType;
import org.apache.poi.EncryptedDocumentException;

//import com.sun.xml.bind.v2.schemagen.xmlschema.List;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;

public class JS_Report_Extraction {

//	static String fixPath="D:\\asd";
//	static File dir = new File(fixPath);
//	static String reportPath = "C:\\\\Users\\\\murtaza.mithaiwala.ZENSOFT\\\\Zensoft_Services_Pvt._Ltd\\\\Sneha_Kanade_-_Philips_Project\\\\2021\\\\Suite_Result\\\\Others_Env\\\\utilitytest";
//	static String reportPathOnLocal = "D:\\asd\\MultipleModelReportAnalysis_test";
	static String reportPath;
	static String reportPathOnLocal;
	static String fixPath;
	static File dir;
	//static String Modalitysheet = "C:\\Users\\murtaza.mithaiwala.ZENSOFT\\Zensoft_Services_Pvt._Ltd\\Sneha_Kanade_-_Philips_Project\\2021\\Suite_Result\\Others_Env\\Modality";
	static String Modalitysheet = "C:\\OneDrive\\Zensoft Services Pvt. Ltd\\PHILIPS - Suite Result\\Others Env\\Modality";	
	static String fixPath1;
	static String filename=null;
	static int excelPrintIterator = 0;
	static int i = 0;
	static String newString = null;
	static int latestCreatedFolder=0;
	static String mostRecentITR_Folder;
	static int itrSubFolderVal=1;
	static int counter =1;
	static String quoteNumber=null;
	static String errorMessage = null;
	static String modelInt;
	static String modelInt2;
	static int subFolderCntr=-1;
	static String stringDate;
	static String TestcaseName;
	static String ExecutionDateandTime;
	static String TCStatus;
	static String TCSeqId;
	private static Element node;
	static Element eElement = (Element) node;
	static Map<String, String> SingleTChildParent = new HashMap<String, String>();
	static Map<String, String> DoubleTChildParent = new HashMap<String, String>();
	static Map<String, List<String>> multiWrenchModel = new HashMap<String, List<String>>();
	static Map<String, String> TCexecutionStatus = new HashMap<String, String>();
	static HashMap<String,String> HashMapQuoteCapt=new HashMap<String,String>();
	static HashMap<String,String> HashMapTcStatus=new HashMap<String,String>();
	static HashMap<String,String> HashMapreportLink=new HashMap<String,String>();
	static Map<String, String> HashMapErrorMessages = new HashMap<String, String>();
	static Map<String, String> ExtraRules = new HashMap<String, String>();
	static Map<String, String> evenAfterRules = new HashMap<String, String>();
	static Map<String, String> ExpectedRules = new HashMap<String, String>();
	static Map<String,String> NotEradicatedRules = new HashMap<String,String>();
	static Map<String,String>HashMapScreenshot=new HashMap<String,String>();
	static Map<String,String>TCcontainsModels=new HashMap<String,String>();
	static Map<String,String>CountryCodeMAp=new HashMap<String,String>();
	static Map<String,String>SeceuencedFailureMap=new HashMap<String,String>();
	static ArrayList<String> array_tcName = new ArrayList<String>();
	static ArrayList<String> array_tcNamefrominfo = new ArrayList<String>();
	static ArrayList<String> array_tcStatus = new ArrayList<String>();
	static ArrayList<String> array_modelNumber = new ArrayList<String>();
	static ArrayList<String> array_TCName = new ArrayList<String>();
	static ArrayList<String> TestcaseNameArrry = new ArrayList<String>();
	static ArrayList<String> FailedTestcaseNameArrry = new ArrayList<String>();
	static String MainReportPath;
	static boolean modelPresent=false;
	static boolean quotePresent=false;
	static boolean errorpresent=false;
	static boolean ExtraRulesPresent=false;
	static boolean evenAfterRulesPresent=false;
	static boolean ExpectedRulesPresent=false;
	static boolean NotEradicatedRulesPresent=false;
	static boolean CountryCodePresent=false;
	static Map<String, String> env;
	static String resultPath="D:\\\\Offline\\\\Reports";
	static String finalExcelReportPath;  
	static int Screenshotcntr = 0;
	static String ExecutionDate;
	static String ExecutionHours;
	static String ExecutionMins;
	static String ExecutionTime;
	static String[] resultDate;
	static String Date;
	static String SuiteName;
	static int p,q,r,s,t;
	static String ExpectedRulesString;
	static String HashMapErrorMessagesString;
	static String ExtraRulesString;
	static String evenAfterRulesString;
	static String NotEradicatedRulesString;
	static int LastSrNoCount;
	static String modelInt2dupli;
	static boolean modelInt2dupliPresent;
	static int FailedCount = 0;
	static String SequencedFailedReasons = "";

	public static void main(String[] args) throws IOException 
	{	
 		fixPath=args[0];
		dir = new File(fixPath);
		reportPathOnLocal=args[1];
		reportPath=args[2];
		reportPath = reportPath.replace('_',' ');
		Modalitysheet  = Modalitysheet.replace('_',' ');
		System.out.println("reportpath"+reportPath);

		for(int k=0;k<args.length;k++)
		{
			System.out.println(args[k]);
		}
		initializeMap("SingleTChildParent");
		initializeMap("DoubleTChildParent");
		initializeMap("MultiWrenchModels");
		latestFolderName();
		subFoldercount();

		Date date = (Date) new java.util.Date();
		SimpleDateFormat DateFor = new SimpleDateFormat("MMMM_dd");
		stringDate = DateFor.format(date);

		//suiteTcStatus();
		errorMessageCapt();
		reportLinkGenration();	
		//quoteNumberCapt();

	}

	public static void initializeMap(String SheetName)  throws IOException{

		String Modalitysheet1;
		Modalitysheet1 = Modalitysheet+".xls";
		File f = new File(Modalitysheet1);
		if (f.exists())                       //if Excel file exist 
		{
			FileInputStream inputStream = new FileInputStream(new File(Modalitysheet1));
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			HSSFSheet sheet = workbook.getSheet(SheetName);  
			int rowCount = sheet.getLastRowNum();
			for(i=0;i <= rowCount; i++) {
				Row row = sheet.getRow(i); 
				String ChildModel= row.getCell(0).toString();  
				if(ChildModel.contains(".0")) {
					ChildModel = ChildModel.replace(".0", "");
				}
				String ParentModel= row.getCell(1).toString(); 
				if(ParentModel.contains(".0")) {
					ParentModel = ParentModel.replace(".0", "");
				}
				String CountryCodeinExcel = null;
				ArrayList<String> ParentAndCountry = new ArrayList<String>();
				if(SheetName.contentEquals("MultiWrenchModels")) {
					CountryCodeinExcel = row.getCell(2).toString(); 
					if(CountryCodeinExcel.contains(".0")) {
						CountryCodeinExcel = CountryCodeinExcel.replace(".0", "");
					}
					ParentAndCountry.add(ParentModel);
					ParentAndCountry.add(CountryCodeinExcel);
				}
				
				if(SheetName.contentEquals("SingleTChildParent"))
					SingleTChildParent.put(ChildModel,ParentModel);
				else if(SheetName.contentEquals("DoubleTChildParent"))
					DoubleTChildParent.put(ChildModel,ParentModel);
				else if(SheetName.contentEquals("MultiWrenchModels")) {
					multiWrenchModel.put(ChildModel,ParentAndCountry);
				}
			}
			workbook.close();
			inputStream.close();
		}
		else {
			System.out.println("Modality/CHildParentFile File Not Fouund");
		}
	
		System.out.println("SingleTChildParent "+SingleTChildParent+" /n "+"DoubleTChildParent "+DoubleTChildParent);

	}

	public static void errorMessageCapt() 
	{
		try {
			String mostRecentITR_Folder=Integer.toString(latestCreatedFolder);
			String itrFolderVal="ITR_"+mostRecentITR_Folder;
			int itrSubFolderVal=1;
			int a;
			int SuiteNameCount = 1;
			int SuiteDateFormatCount = 1;

			for(int itr=0;itr<subFolderCntr;itr++)
			{
				File myObj = new File(dir+"\\"+itrFolderVal+"\\"+itrSubFolderVal+"\\ITR_1\\InfoLog.html");
				itrSubFolderVal++;
				int ModelNumberCount = 1;
				int QuoteNumberCount = 1;
				Scanner myReader = new Scanner(myObj);
				Scanner MultiWrenchmyReader = new Scanner(myObj);
				ArrayList<String> array_quoteNumber = new ArrayList<String>();
				ArrayList<String> errorMessagearry = new ArrayList<String>();
				ArrayList<String> extrarulesarry = new ArrayList<String>();
				ArrayList<String> rulesEvenAfter = new ArrayList<String>();
				ArrayList<String> Expectedrulesarry = new ArrayList<String>();
				ArrayList<String> NotEradicatedrulesarry = new ArrayList<String>();
				String noRulesString = "-";
				String noErrorMessagearryString = "No error";
				int FirstErrorMessage = 0;
				LinkedList<String> fivePrevLines = new LinkedList<>();
				ExpectedRulesString="";
				HashMapErrorMessagesString="";
				ExtraRulesString="";
				evenAfterRulesString="";
				NotEradicatedRulesString="";
				Boolean MultiWrenchPresent = false;
				String CountryCode = "";
				p =0;
				q=0;
				r=0;
				s=0;
				t=0;
				String ModelNumberInfo ="NA";
				quoteNumber="NA";
				String ProductCodeNum = "";
				int errorOptionCnt = 0;
				boolean insideChild = false;
				boolean childEnds = false;
				boolean childStarts = false;
				String ChildBundle = "";
				String secondProduct = "";
				String firstProduct = "";
//				while (MultiWrenchmyReader.hasNextLine()) 
//				{
//					String data = MultiWrenchmyReader.nextLine();
//					if(data.contains("Multiple Wrench icon validation on cart page"))
//					{
//						MultiWrenchPresent = true;	
//					}
//				}
				while (myReader.hasNextLine()) 
				{
					String data = myReader.nextLine();
					fivePrevLines.add(data);
					if (fivePrevLines.size() > 10) {
						fivePrevLines.removeFirst();
					}
					
//					if(QuoteNumberCount == 1) {					
//						if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ") && data.contains(" is stored successfully in the key 'QuoteNumberToCLone'</span>"))
//						{  
//							String rmWord="<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ";
//							quoteNumber=data.replaceAll(rmWord, "");
//							quoteNumber=quoteNumber.substring(10,20);
//							quotePresent=true;
//							QuoteNumberCount++;
//						}
//					}

					if(QuoteNumberCount == 1) {					
						if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreText    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ")&& data.contains(" is stored successfully in the key 'QuoteNumber'</span>"))
						{  
						   String rmWord="<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreText    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ";
							quoteNumber=data.replaceAll(rmWord, "");
							quoteNumber=quoteNumber.substring(10,20);
							quotePresent=true;
							QuoteNumberCount++;
						}
					}			

					if(ModelNumberCount == 1) {
						if(data.contains("Step Itinerary =>"))
						{
							String rmword = "<span name='Message' class='log-span step-log'><span class='log-label'>Step Itinerary => ";
							String SuiteNameLine = data.replaceAll(rmword,"");
							String[] SuiteNamearay = SuiteNameLine.split(" >> ");
							System.out.println("SuiteNamearay "+SuiteNamearay[2]);
							String TcNamearray[] = SuiteNamearay[2].split("_");
							modelInt2 = TcNamearray[0];
							
							
//							if(MultiWrenchPresent == true) {
//								modelInt2 = TcNamearray[0]+" & "+TcNamearray[1];
//								secondProduct = TcNamearray[1];
//								firstProduct = TcNamearray[0];
//								SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+firstProduct;
//							}
							System.out.println("ModelNumberInfo: "+modelInt2);
							ModelNumberCount++;
						}
					} 

					if(SuiteNameCount == 1) {
						for(i = 0; i < 15; i++) {
							if(data.contains("Step Itinerary =>"))
							{
								String rmword = "<span name='Message' class='log-span step-log'><span class='log-label'>Step Itinerary => ";
								String SuiteNameLine = data.replaceAll(rmword,"");
								String[] SuiteNamearay = SuiteNameLine.split(" >> ");
								SuiteName = SuiteNamearay[0];	
								System.out.println("SuiteName in info: "+SuiteName);
								break;
							}
							data = myReader.nextLine();
						}
						for(i = 0; i < 15; i++) {
							if(data.contains("Execution Begin at: "))
							{
								String rmword = "<span name='Message' class='log-span'><span class='log-label'>Action Execution Begin at: </span>";
								String SuiteTimeLine = data.replaceAll(rmword,"");
								String[] SuiteTimearay = SuiteTimeLine.split("</span>");
								ExecutionDateandTime = SuiteTimearay[0];	
								ExecutionDate = ExecutionDateandTime.substring(0,10);
								ExecutionHours= ExecutionDateandTime.substring(11,13);
								ExecutionMins = ExecutionDateandTime.substring(14,16);
								ExecutionTime = ExecutionHours+" Hrs "+ExecutionMins+" mins";
								System.out.println(ExecutionDate);
								System.out.println("execution time from info"+ExecutionTime);
								break;
							}
							data = myReader.nextLine();
						}
						System.out.println("SuiteNameCount "+SuiteNameCount);
					}
					SuiteNameCount++;


					if(data.contains("ModelNumber"))
					{

						String rmWord= "<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ";
						String modelNumber=data.replaceAll(rmWord, "");
						String[] ModelString = modelNumber.split("'", 0);
						modelInt = ModelString[1];
						if(modelInt.contains(" ")) {
							String[] ModelNumberString2 = modelInt.split(" ", 0);
							modelInt = ModelNumberString2[0];
						}
						System.out.println(modelInt+": modelInt from TestData for SS"); 
					
					}
					
					if(data.contains("The Country Code"))
					{

						String rmWord= "<span name='Message' class='log-span'><span class='log-label'>Action: </span> PrintMessageInReport    <span class='log-label'> Status: </span>  passed    <span class='log-label'> Message: </span>";
						String CountryCodestr =data.replaceAll(rmWord, "");
						String[] CountryCodeArray = CountryCodestr.split("'", 0);
						CountryCode = CountryCodeArray[1];
						CountryCodePresent = true;		
						
						Iterator <String> iterator = multiWrenchModel.keySet().iterator();
						while(iterator.hasNext())  
						{   
							String key= iterator.next();
							
							if(key.contentEquals(modelInt2)) {	
								List<String> CountryFromExcel = new ArrayList<>();
								CountryFromExcel = multiWrenchModel.get(key);
								if(CountryFromExcel.get(1).contentEquals(CountryCode)) {
								MultiWrenchPresent = true;	
								modelInt2 = modelInt2+" & "+CountryFromExcel.get(0);	
								secondProduct = CountryFromExcel.get(0);
								firstProduct = modelInt2;
								SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+firstProduct;
								}
							}
						}
					}
					
				
					
					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
								//SequencedFailedReasons = SequencedFailedReasons.repl
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					//if(data.contains("</span>  Failed") || data.contains("</span>  failed")||data.contains("</span> Failed"))
					if(data.contains("</span>  Failed") || data.contains("</span>  failed")||data.contains("</span> Failed") || data.contains("Failed") || data.contains("failed"))
					{ 
						if(FirstErrorMessage == 0 ) {
							FirstErrorMessage++;
							errorMessagearry.add("Failed: ");
							//SequencedFailedReasons = SequencedFailedReasons+"Failed: ";
						}

						String[] errorMessagespltarry = data.split("</span>",0);
						a = errorMessagespltarry.length;	
						if(errorMessagespltarry[a-1].contains("not found on Catlog page")||errorMessagespltarry[a-1].contains("not found on Configuration page")||errorMessagespltarry[a-1].contains("Spinning wheel")||errorMessagespltarry[a-1].contains("Greyed Out")||errorMessagespltarry[a-1].contains("1/1 match")||errorMessagespltarry[a-1].contains("0/0 Match")||errorMessagespltarry[a-1].contains("Auto selected")||errorMessagespltarry[a-1].contains("Either Change")||errorMessagespltarry[a-1].contains("Go - to - Pricing")||errorMessagespltarry[a-1].contains("Wrenchicon")||errorMessagespltarry[a-1].contains("Actual string") || errorMessagespltarry[a-1].contains("Product Description on catalog")||errorMessagespltarry[a-1].contains("Product Code on catalog")||errorMessagespltarry[a-1].contains("NPO")||errorMessagespltarry[a-1].contains("not available"))
						{
							errorMessagearry.add(errorMessagespltarry[a-1]);
							System.out.println("SequencedFailedReasons : "+SequencedFailedReasons+" /n"+modelInt2+" Before if outside while ChildBundle: '"+ChildBundle+"' /nProductCodeNum: '"+ProductCodeNum+"'");

							if((errorMessagespltarry[a-1].contains("Wrenchicon Shows pending configuration")||errorMessagespltarry[a-1].contains("Go - to - Pricing Disabled.")) && SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								System.out.println("SequencedFailedReasons : "+SequencedFailedReasons+" /n"+modelInt2+" inside if outside while ChildBundle: '"+ChildBundle+"' /nProductCodeNum: '"+ProductCodeNum+"'");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
								if(errorMessagespltarry[a-1].contains("Wrenchicon Shows pending configuration") && SequencedFailedReasons.endsWith("Product/Model :"+secondProduct))
								{
									SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+secondProduct,"");
									System.out.println(modelInt2+"inside if if secondProduct: "+secondProduct);
								}
								SequencedFailedReasons = SequencedFailedReasons+"\n"+errorMessagespltarry[a-1]+"\n";
								System.out.println(modelInt2+"inside if");
							}
							else 
							{
								System.out.println(modelInt2+"inside else secondProduct: "+secondProduct);
								SequencedFailedReasons = SequencedFailedReasons+errorMessagespltarry[a-1]+"\n";
							}
						}	        		
						errorpresent = true;
					}


					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					if(data.contains("The follwing value(s) are in expected Data, but not available on UI - "))
					{
						if(errorOptionCnt == 0) {
							SequencedFailedReasons = SequencedFailedReasons+"Before selecting Options Expected Rules Not on UI:"+"\n";
						}
						else {
							SequencedFailedReasons = SequencedFailedReasons+" Rules Not on UI:"+"\n";
						}
						data = myReader.nextLine();
						int ExpectedErrorCount = 1;
						do {		    			
							String[] expectedrulesarry = data.split("'>"+ExpectedErrorCount+".",0);
							int b = expectedrulesarry.length;
							String[] qaz = expectedrulesarry[b-1].split("</",0);
							Expectedrulesarry.add(qaz[0]);
							SequencedFailedReasons = SequencedFailedReasons+qaz[0]+"\n";
							ExpectedErrorCount++;
							data = myReader.nextLine();
							ExpectedRulesPresent  = true;
						}while(data.contains(ExpectedErrorCount+"."));
					}


					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					if(data.contains("The follwing value(s) are in expected Data, which should not be available on UI, but those are available on UI -"))
					{

							SequencedFailedReasons = SequencedFailedReasons+" Rules Not Eradicated:"+"\n";
						
						data = myReader.nextLine();
						int NotEradicatedErrorCount = 1;
						do {
							String[] expectedrulesarry = data.split("'>"+NotEradicatedErrorCount+".",0);
							int b = expectedrulesarry.length;
							String[] qaz = expectedrulesarry[b-1].split("</",0);
							NotEradicatedrulesarry.add(qaz[0]);
							SequencedFailedReasons = SequencedFailedReasons+qaz[0]+"\n";
							NotEradicatedErrorCount++;
							data = myReader.nextLine();
							NotEradicatedRulesPresent  = true;
						}while(data.contains(NotEradicatedErrorCount+"."));
					}


					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					if(data.contains("The follwing value(s) are on UI, but not available in Test Data - "))
					{
						if(errorOptionCnt == 0) {
							SequencedFailedReasons = SequencedFailedReasons+"Before selecting Options Extra Rules on UI:"+"\n";
						}else {
							SequencedFailedReasons = SequencedFailedReasons+" Extra Rules on UI:"+"\n";
						}
						data = myReader.nextLine();
						int ErrorCount = 1;
						do {
							String[] rulesarry = data.split("'>"+ErrorCount+".",0);
							int b = rulesarry.length;
							String[] qaz = rulesarry[b-1].split("</",0);			   
							extrarulesarry.add(qaz[0]);
							SequencedFailedReasons = SequencedFailedReasons+qaz[0]+"\n";
							ErrorCount++;
							ExtraRulesPresent = true;
							data = myReader.nextLine();
						}while(data.contains(ErrorCount+"."));
					}
					

					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					if(data.contains("Element is not visible for comparison with expected data"))
					{
						if(errorOptionCnt == 0) {
							SequencedFailedReasons = SequencedFailedReasons+"Before selecting Options Expected Rules Not on UI:"+"\n";
						}
						else {
							SequencedFailedReasons = SequencedFailedReasons+" Rules Not on UI:"+"\n";
						}
						int ExpectedErrorCount = 1;
						for (String preline : fivePrevLines) {
							if(preline.contains(ExpectedErrorCount+". ")){
								//							if(preline.contains("The values fetched from the Test Data - ")){
								//								preline = myReader.nextLine();
								String[] expectedrulesarry = preline.split("'>"+ExpectedErrorCount+".",0);
								int b = expectedrulesarry.length;
								String[] qaz = expectedrulesarry[b-1].split("</",0);
								Expectedrulesarry.add(qaz[0]);
								SequencedFailedReasons = SequencedFailedReasons+qaz[0]+"\n";
								ExpectedErrorCount++;
							}
							ExpectedRulesPresent  = true;
						}
					} 


					if(data.contains("<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  The data ") && data.contains("is stored successfully in the key 'ProductCode'</span>")) 
					{						
						String ProductCodeArry[] = data.split("'",0);
						int ProductCodesize = ProductCodeArry.length;
						String ProductCode = ProductCodeArry[ProductCodesize - 4];
						if(insideChild == false) {
							for(int i = 0; i <= 45; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickWrenchIcon'")) {
								childStarts = true;
								insideChild = true;
								break;
								}
							}
						}
						
						if(!ProductCode.contentEquals("")) 
						{
							errorOptionCnt++;
							if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
							{
								//System.out.println("endsWith");
								SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							}
							ProductCodeNum = ProductCode;
							if(childStarts == true) {
								ChildBundle = "Inside Child Bundle-> "+ProductCode+"-> ";	
								childStarts = false;
							}
							if(childEnds == true) {
								ChildBundle = "";	
								childEnds = false;
							}
							SequencedFailedReasons = SequencedFailedReasons+"\n"+ChildBundle+"Option-> "+ProductCode+" Failure Reasons: \n";
						}
						
						if(insideChild == true) {
							for(int i = 0; i <= 54; i++) {
								data = myReader.nextLine();
								if(data.contains("The data 'true' is stored successfully in the key 'clickMainBundleLink'")) {
								childEnds = true;
								insideChild = false;
								break;
								}
							}
						}
						
						data = myReader.nextLine();
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
					if(data.contains("ErrorNotExpect")){
						for(i = 0; i < 3; i++) {
							data = myReader.nextLine();
						}
						int ErrorCount = 1;
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n After All Option Selection following Errors: \n";
						do {
							String[] rulesarry = data.split("'>"+ErrorCount+".",0);
							int b = rulesarry.length;
							String[] qaz = rulesarry[b-1].split("</",0);
							rulesEvenAfter.add(qaz[0]);
							SequencedFailedReasons = SequencedFailedReasons+qaz[0]+"\n";
							ErrorCount++;
							evenAfterRulesPresent = true;
							data = myReader.nextLine();
						}while(data.contains(ErrorCount+"."));
					}
					
					if(MultiWrenchPresent==true && data.contains(" The value '"+secondProduct)) {
						if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
						{
							SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
							if(SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
								SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
							}
						}
						SequencedFailedReasons = SequencedFailedReasons+"\n"+"Product/Model :"+secondProduct;
						data = myReader.nextLine();
					}
					
				} 
//				System.out.println(modelInt2+" Before if outside while ChildBundle: '"+ChildBundle+"' ProductCodeNum: '"+ProductCodeNum+"'");
//				if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
//				{
//					System.out.println(modelInt2+" if outside while ChildBundle: '"+ChildBundle+"' ProductCodeNum: '"+ProductCodeNum+"'");
//					SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
//				}
				
				if(SuiteDateFormatCount == 1) {
					SuiteDateFormatCount++;
					if(ExecutionDate.endsWith(" ")) {
						resultDate =  ExecutionDate.split(" ");
						ExecutionDate = resultDate[0];
					}

					if(ExecutionDate.startsWith("202")) {
						String[] Datesarry = ExecutionDate.split("-");
						System.out.println(Datesarry[1]);
						if(Datesarry[1].equals("01")) {
							Date = "Jan_"+Datesarry[2];
						}
						if(Datesarry[1].equals("02")) {
							Date = "Feb_"+Datesarry[2];
						}
						if(Datesarry[1].equals("03")) {
							Date = "Mar_"+Datesarry[2];
						}
						if(Datesarry[1].equals("04")) {
							Date = "April_"+Datesarry[2];
						}
						if(Datesarry[1].equals("05")) {
							Date = "May_"+Datesarry[2];
						}
						if(Datesarry[1].equals("06")) {
							Date = "June_"+Datesarry[2];
						}							      
						if(Datesarry[1].equals("07")) {
							Date = "July_"+Datesarry[2];
						}
						if(Datesarry[1].equals("08")) {
							Date = "Aug_"+Datesarry[2];
						}
						if(Datesarry[1].equals("09")) {
							Date = "Sept_"+Datesarry[2];
						}
						if(Datesarry[1].equals("10")) {
							Date = "October_"+Datesarry[2];
						}
						if(Datesarry[1].equals("11")) {
							Date = "Nov_"+Datesarry[2];
						}
						if(Datesarry[1].equals("12")) {
							Date = "Dec_"+Datesarry[2];
						}
					}
					System.out.println("Date from info "+Date);
				}    



				myReader.close();

				ArrayList<String> newExpectedrulesarry = new ArrayList<String>(); 
				ArrayList<String> newextrarulesarry = new ArrayList<String>(); 
				ArrayList<String> newNotEradicatedrulesarry = new ArrayList<String>(); 
				ArrayList<String> newrulesEvenAfter = new ArrayList<String>(); 

				for (String element : rulesEvenAfter) { 

					if (!newrulesEvenAfter.contains(element)) { 

						newrulesEvenAfter.add(element);

					} 
				}

				for (String element : NotEradicatedrulesarry) { 

					if (!newNotEradicatedrulesarry.contains(element)) { 

						newNotEradicatedrulesarry.add(element);

					} 
				}
				for (String element : Expectedrulesarry) { 

					if (!newExpectedrulesarry.contains(element)) { 

						newExpectedrulesarry.add(element);

					} 
				}
				for (String element : extrarulesarry) { 

					if (!newextrarulesarry.contains(element)) { 

						newextrarulesarry.add(element);

					} 
				}

				List<String> intersection = new ArrayList<String>(newNotEradicatedrulesarry);
				intersection.retainAll(newextrarulesarry);
				newextrarulesarry.removeAll(intersection);
				if(newextrarulesarry.isEmpty()) {
					ExtraRulesPresent = false;
				}
				int aq = 1;
				for(String i : newExpectedrulesarry) {
					if(aq == 1)
						ExpectedRulesString = ExpectedRulesString+aq+")"+i;
					else
						ExpectedRulesString = ExpectedRulesString+"\n"+aq+")"+i;
					aq++;
				}

				int aw = 1;
				for(String i : newNotEradicatedrulesarry) {
					if(aw == 1)
						NotEradicatedRulesString = NotEradicatedRulesString+aw+")"+i;
					else
						NotEradicatedRulesString = NotEradicatedRulesString+"\n"+aw+")"+i;
					aw++;
				}

				int ae = 1;
				for(String i : newextrarulesarry) {
					if(ae == 1)
						ExtraRulesString = ExtraRulesString+ae+")"+i;
					else
						ExtraRulesString = ExtraRulesString+"\n"+ae+")"+i;
					ae++;
				}

				int ar = 1;
				for(String i : newrulesEvenAfter) {
					if(ar == 1)
						evenAfterRulesString = evenAfterRulesString+ar+")"+i;
					else
						evenAfterRulesString = evenAfterRulesString+"\n"+ar+")"+i;
					ar++;
				}
				System.out.println("errorMessagearry: "+errorMessagearry);
				int at = 1;
				for(String i : errorMessagearry) {
					if(at == 1 || at == 2)
						HashMapErrorMessagesString = HashMapErrorMessagesString+i;
					else if(i.contains("Wrenchicon Shows pending configuration.")) {
						System.out.println("in wrench if");
						HashMapErrorMessagesString = HashMapErrorMessagesString+i;
					}
					else
						HashMapErrorMessagesString = HashMapErrorMessagesString+"\n"+i;
					at++;
				}

				System.out.println("HashMapErrorMessagesString :"+HashMapErrorMessagesString);
				System.out.println("Expected Rules String"+ExpectedRulesString);
				System.out.println( " intersection -"+intersection);
				System.out.println("Not Eradicated - "+newNotEradicatedrulesarry);
				System.out.println("Expected rules - "+newExpectedrulesarry);
				System.out.println("extra rules - "+newextrarulesarry);
				System.out.println("even after - "+newrulesEvenAfter);

				modelInt2 = modelInt2+"-"+CountryCode;

				Iterator <String> it = SingleTChildParent.keySet().iterator();

				while(it.hasNext())  
				{   
					String key= it.next();
					if(key.contentEquals(modelInt2)) {
						String modelInt2array[] = modelInt2.split("-");
						modelInt2 = modelInt2array[0]+"_"+SingleTChildParent.get(key)+"-"+modelInt2array[1];					
					}
				}	

				Iterator <String> it1 = DoubleTChildParent.keySet().iterator();

				while(it1.hasNext())  
				{   
					String key= it1.next();
					if(DoubleTChildParent.get(key).contentEquals(modelInt2)) {
						modelInt2dupli = key+"_"+modelInt2;
						modelInt2dupliPresent = true;
					}
				}
				if(modelInt2dupliPresent == true) {
					array_tcNamefrominfo.add(modelInt2dupli);
				}
				array_tcNamefrominfo.add(modelInt2);

				if(SequencedFailedReasons.endsWith("Failure Reasons: \n"))
				{
					SequencedFailedReasons = SequencedFailedReasons.replace("\n"+ChildBundle+"Option-> "+ProductCodeNum+" Failure Reasons: \n", "");
				}
				
				if(MultiWrenchPresent==true && SequencedFailedReasons.endsWith("Product/Model :"+secondProduct)) {
					SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+secondProduct,"");
				}
				if(MultiWrenchPresent==true && SequencedFailedReasons.endsWith("Product/Model :"+firstProduct)) {
					SequencedFailedReasons = SequencedFailedReasons.replace("Product/Model :"+firstProduct,"");
				}
				
				if(errorpresent == false && ExpectedRulesPresent == false && NotEradicatedRulesPresent == false && ExtraRulesPresent == false && evenAfterRulesPresent == false)
				{ 
					TCexecutionStatus.put(modelInt2, "0");
					if(modelInt2dupliPresent == true) {
						TCexecutionStatus.put(modelInt2dupli, "0");
					}
					
					SeceuencedFailureMap.put(modelInt2, "-");
					if(modelInt2dupliPresent == true) {
						SeceuencedFailureMap.put(modelInt2dupli, "-");
					}
					
				}
				else {
					FailedCount++;
					System.out.println("FailedCount "+FailedCount);
					TCexecutionStatus.put(modelInt2, "1");
					if(modelInt2dupliPresent == true) {
						TCexecutionStatus.put(modelInt2dupli, "1");
					}
					if(modelInt2dupliPresent == true) {
						SeceuencedFailureMap.put(modelInt2dupli, "Failed: "+SequencedFailedReasons);
					}
					
					SeceuencedFailureMap.put(modelInt2, "Failed: "+SequencedFailedReasons);
				
					System.out.println("fail");
				} 
				
				SequencedFailedReasons = "";

				HashMapQuoteCapt.put(modelInt2, quoteNumber);
				if(modelInt2dupliPresent == true) {
					HashMapQuoteCapt.put(modelInt2dupli, quoteNumber);
				}

				if(errorpresent == true) {
					System.out.println(modelInt2);
					if(modelInt2dupliPresent == true) {
						HashMapErrorMessages.put(modelInt2dupli, HashMapErrorMessagesString);
					}
					HashMapErrorMessages.put(modelInt2,HashMapErrorMessagesString);
				}
				else {
					System.out.println(modelInt2);
					if(modelInt2dupliPresent == true) {
						HashMapErrorMessages.put(modelInt2dupli,noErrorMessagearryString);
					}
					HashMapErrorMessages.put(modelInt2,noErrorMessagearryString);
				}

				if(ExtraRulesPresent == true) {
					ExtraRules.put(modelInt2, ExtraRulesString);
					if(modelInt2dupliPresent == true) {
						ExtraRules.put(modelInt2dupli, ExtraRulesString);
					}
					System.out.println("modelInt2 - ExtraRulesString"+modelInt2+ExtraRulesString);
				}
				else
				{
					ExtraRules.put(modelInt2,"-");
					if(modelInt2dupliPresent == true) {
						ExtraRules.put(modelInt2dupli,"-");
					}
				}
				if(ExpectedRulesPresent  == true) {
					ExpectedRules.put(modelInt2, ExpectedRulesString);
					if(modelInt2dupliPresent == true) {
						ExpectedRules.put(modelInt2dupli, ExpectedRulesString);
					}
					System.out.println("modelInt2 - ExpectedRulesString"+modelInt2+ExpectedRulesString);
				}
				else
				{
					ExpectedRules.put(modelInt2,"-");
					if(modelInt2dupliPresent == true) {
						ExpectedRules.put(modelInt2dupli,"-");
					}
				}
				if(NotEradicatedRulesPresent  == true) {
					NotEradicatedRules.put(modelInt2,NotEradicatedRulesString );
					if(modelInt2dupliPresent == true) {
						NotEradicatedRules.put(modelInt2dupli,NotEradicatedRulesString );
					}
					System.out.println("modelInt2 - NotEradicatedRulesString"+modelInt2+NotEradicatedRulesString);
				}
				else {
					NotEradicatedRules.put(modelInt2,"-");
					if(modelInt2dupliPresent == true) {
						NotEradicatedRules.put(modelInt2dupli,"-");
					}					
				}
				if(evenAfterRulesPresent  == true && newExpectedrulesarry.isEmpty() && newextrarulesarry.isEmpty() && newNotEradicatedrulesarry.isEmpty()) {
					evenAfterRules.put(modelInt2,evenAfterRulesString );
					if(modelInt2dupliPresent == true) {
						evenAfterRules.put(modelInt2dupli,evenAfterRulesString );
					}
					System.out.println("modelInt2 - evenAfterRulesString"+modelInt2+evenAfterRulesString);
				}
				else {
					evenAfterRules.put(modelInt2,"-");
					if(modelInt2dupliPresent == true) {
						evenAfterRules.put(modelInt2dupli,"-");
					}
				}
				if(CountryCodePresent = true) {
					CountryCodeMAp.put(modelInt2, CountryCode);
					if(modelInt2dupliPresent == true) {
						CountryCodeMAp.put(modelInt2dupli, CountryCode);
					}
				}
				else {
					CountryCodeMAp.put(modelInt2, "NA");
					if(modelInt2dupliPresent == true) {
						CountryCodeMAp.put(modelInt2dupli, "NA");
					}
				}
				if(errorpresent == false && ExtraRulesPresent == false && ExpectedRulesPresent == false && NotEradicatedRulesPresent  == false && evenAfterRulesPresent== false) {
					errorScreenshot("---");
				}
				else {
					//String ModelNumber = String.valueOf(modelInt);
					errorScreenshot(modelInt);
//					if(modelInt2dupliPresent == true) {
//						ModelNumber = String.valueOf(modelInt2dupli);
//						errorScreenshot(ModelNumber);
//					}
				}
				NotEradicatedRulesPresent  = false;
				errorpresent = false;
				ExtraRulesPresent = false;
				ExpectedRulesPresent = false;
				CountryCodePresent = false;
				evenAfterRulesPresent = false;
				modelInt2dupliPresent = false;
				modelInt2dupli = "NA";
				noRulesString = "-";
				noErrorMessagearryString = "No error";
				System.out.println(HashMapErrorMessages.get(modelInt2));
			}
			System.out.println("NotEradicatedRules :"+NotEradicatedRules);
			System.out.println("evenAfterRules :"+evenAfterRules);
			System.out.println("ExpectedRules :"+ExpectedRules);
			System.out.println("ExtraRules :"+ExtraRules);	
			System.out.println("SeceuencedFailureMap "+SeceuencedFailureMap);
/*
			for(excelPrintIterator = 1;excelPrintIterator < 3;excelPrintIterator++) {
				if(excelPrintIterator == 1) {
					printinExcel(reportPathOnLocal);
					System.out.println("Excel written successfully..on Local");
				}
				else {
					printinExcel(reportPath);
					System.out.println("Excel written successfully..on Drive");
				}
			}
			
*/
		}

		catch (FileNotFoundException e) 
		{
			System.out.println("An error occurred.");
			e.printStackTrace();
		}
	}

	public static void errorScreenshot(String folderName) {
		try {
			if(Screenshotcntr == 0) 
			{		    Screenshotcntr++;
			ExecutionDate = ExecutionDate.replace('-','_');
			System.out.println("ExecutionDate in ss func. "+ExecutionDate);
			}
			if(folderName == "---") {
				HashMapScreenshot.put(modelInt2,"---");
				if(modelInt2dupliPresent == true) {
					HashMapScreenshot.put(modelInt2dupli,"---");
				}
			}
			else {	
				String sreenShotFolderPath = "C:\\Screenshot_FailedTC\\"+folderName+"_"+ExecutionDate;
				Path folderpath=Paths.get(sreenShotFolderPath);
				boolean fileExist = Files.exists(folderpath);

				if(fileExist) {
					HashMapScreenshot.put(modelInt2,sreenShotFolderPath);
					if(modelInt2dupliPresent == true) {
						HashMapScreenshot.put(modelInt2dupli,sreenShotFolderPath);
					}
				}
				else {
					HashMapScreenshot.put(modelInt2,"-/-");
					if(modelInt2dupliPresent == true) {
						HashMapScreenshot.put(modelInt2dupli,"-/-");
					}
				}
			}

			System.out.println(HashMapScreenshot);
		}
		catch (Exception e)
		{
			System.out.println(e);
		}
	}

	public static void quoteNumberCapt() throws EncryptedDocumentException, IOException
	{
		try 
		{
			ArrayList<String> array_quoteNumber = new ArrayList<String>();
			String mostRecentITR_Folder=Integer.toString(latestCreatedFolder);
			String itrFolderVal="ITR_"+mostRecentITR_Folder;
			int itrSubFolderVal=1;

			for(int itr=0;itr<subFolderCntr;itr++)
			{
				File myObj = new File(dir+"\\"+itrFolderVal+"\\"+itrSubFolderVal+"\\ITR_1\\InfoLog.html");
				itrSubFolderVal++;
				Scanner myReader = new Scanner(myObj);


				while (myReader.hasNextLine()) 
				{

					String data = myReader.nextLine();
					quoteNumber=null;
					if(data.contains("Q-000"))
					{  
						String rmWord="<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreText    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ";
						quoteNumber=data.replaceAll(rmWord, "");
						quoteNumber=quoteNumber.substring(10,20);
						quotePresent=true;
					}
					if(quotePresent==false)
					{
						quoteNumber="*****";
						array_quoteNumber.add(quoteNumber);
						quotePresent=true;
					}

					if(data.contains("ModelNumber"))
					{

						String rmWord= "<span name='Message' class='log-span'><span class='log-label'>Action: </span> StoreVariable    <span class='log-label'> Status: </span>  Passed    <span class='log-label'> Message: </span>  ";
						String modelNumber=data.replaceAll(rmWord, "");
						String[] ModelString = modelNumber.split("'", 0);
						modelInt = ModelString[1];
						if(modelInt.contains(" ")) {
							String[] ModelNumberString2 = modelInt.split(" ", 0);
							modelInt = ModelNumberString2[0];
						}
						System.out.println(modelInt+": modelInt"); 
						modelPresent=true;
					}
					if(quotePresent==modelPresent)
					{			        	
						array_modelNumber.add(modelInt);
						array_quoteNumber.add(quoteNumber);
					}

				}
				myReader.close();
			} 
			int totalModelNumCount=array_modelNumber.size();
			for(int i=0;i<totalModelNumCount;i++)
			{		    		
				if(array_modelNumber.get(i)!=null && array_quoteNumber.get(i)!=null)
				{
					HashMapQuoteCapt.put(array_modelNumber.get(i), array_quoteNumber.get(i));
				}

			}

			for(excelPrintIterator = 1;excelPrintIterator < 3;excelPrintIterator++) {
				if(excelPrintIterator == 1) {
					printinExcel(reportPathOnLocal);
					System.out.println("Excel written successfully..on Local");
				}
				else {
					//printinExcel(reportPath);
					System.out.println("Excel written successfully..on Drive");
				}
			}


		} 
		catch (FileNotFoundException e) 
		{
			System.out.println("An error occurred.");
			e.printStackTrace();
		}
	}

	public static void latestFolderName()
	{
		System.out.println("Latefol:"+dir);
		try
		{
			FilenameFilter filter = new FilenameFilter() 
			{
				public boolean accept (File dir, String name) 
				{ 
					return name.startsWith("ITR");
				} 
			}; 

			String[] filenameSotreArray = dir.list(filter);
			int size=filenameSotreArray.length;
			int arr[]=new int[size];
			//System.out.println(size);
			if (filenameSotreArray == null)
			{
				System.out.println("Either dir does not exist or is not a directory"); 
			} 

			else
			{ 
				for ( i = 0; i< filenameSotreArray.length; i++)
				{
					filename = filenameSotreArray[i];
					//strLen=filename.length();

					String rmWord="ITR_";
					newString=filename.replaceAll(rmWord, "");


					//int j=0;
					arr[i]=Integer.parseInt(newString);
					///System.out.println(arr[i]);
				}	 
				latestCreatedFolder=arr[0];
				for ( i = 0; i< arr.length; i++)
				{
					if(arr[i] > latestCreatedFolder)  
						latestCreatedFolder = arr[i];
				}
				//System.out.println("#####Most Recently created folder"+latestCreatedFolder);



			}
			// int filenameInt=Integer.parseInt(filename.substring(2,5));
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

	public static void subFoldercount()
	{
		try
		{
			String mostRecentITR_Folder=Integer.toString(latestCreatedFolder);
			String itrFolderVal="ITR_"+mostRecentITR_Folder;
			String abcd=dir+"\\"+itrFolderVal;
			File dir1 = new File(abcd);

			String[] fileNameArray = dir1.list();
			int size=fileNameArray.length;

			if (fileNameArray == null) 
			{
				System.out.println("Either dir does not exist or is not a directory"); 
			} 
			else
			{ 
				subFolderCntr=-1;
				for ( i = 0; i< size; i++)
				{
					filename = fileNameArray[i];
					int strlen=filename.length();
					if(strlen==1||strlen==2||strlen==3)
					{
						subFolderCntr++;
						System.out.println("Inner Data"+filename);
					}

				}
				System.out.println("Subfolder value"+subFolderCntr);	  
			}  
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	public static void suiteTcStatus()
	{
		String mostRecentITR_Folder=Integer.toString(latestCreatedFolder);
		String itrFolderVal="ITR_"+mostRecentITR_Folder;
		try {

			MainReportPath = dir +"\\"+itrFolderVal+"\\"+"SummaryReport.xml";								
			File file = new File(MainReportPath);
			if(file.exists()) {
				Scanner myReader = new Scanner(file);
				while (myReader.hasNextLine()) 
				{
					String data = myReader.nextLine();				
					if(data.contains(" Name=")) 
					{
						String[] Name = data.split(" Name=\"");
						String[] SuiteNamearry = Name[1].split("\" Schedule");
						SuiteName = SuiteNamearry[0];		
					}

					if(data.contains("<Suite EndTime=")) 
					{
						String rmword = "    <Suite EndTime=";
						ExecutionDateandTime = data.replaceAll(rmword,"");
						ExecutionDate = ExecutionDateandTime.substring(1,11);
						ExecutionHours= ExecutionDateandTime.substring(11,13);
						ExecutionMins = ExecutionDateandTime.substring(14,16);
						ExecutionTime = ExecutionHours+" Hrs "+ExecutionMins+" mins";
						//ExecutionDate = ExecutionDate.replace('-','_');
						System.out.println(ExecutionDate);
						System.out.println("execution time "+ExecutionTime);

					}
				}
				myReader.close();
				if(ExecutionDate.endsWith(" ")) {
					resultDate =  ExecutionDate.split(" ");
					ExecutionDate = resultDate[0];
				}
				/*
			if(ExecutionDate.startsWith("202")) {
				String[] Datesarry = ExecutionDate.split("-");
				System.out.println(Datesarry[1]);
				if(Datesarry[1].equals("01")) {
					Date = "Jan_"+Datesarry[2];
				}
				if(Datesarry[1].equals("02")) {
					Date = "Feb_"+Datesarry[2];
				}
				if(Datesarry[1].equals("03")) {
					Date = "Mar_"+Datesarry[2];
				}
				if(Datesarry[1].equals("04")) {
					Date = "April_"+Datesarry[2];
				}
				if(Datesarry[1].equals("05")) {
					Date = "May_"+Datesarry[2];
				}
			
			}
			System.out.println(Date);
				 */
				DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();

				DocumentBuilder db = dbf.newDocumentBuilder();
				Document doc = db.parse(file);
				doc.getDocumentElement().normalize();					
				NodeList nodeList = doc.getElementsByTagName("TC");
				int Total_TC_Count = nodeList.getLength();				
				for (int itr = 0; itr < Total_TC_Count; itr++) 
				{
					Node node = nodeList.item(itr);								
					if (node.getNodeType() == Node.ELEMENT_NODE) 
					{
						Element eElement = (Element) node;
						TestcaseName = eElement.getAttribute("TCName");
						array_TCName.add(TestcaseName);
						String[] ModelString = TestcaseName.split("_", 0);
						String ModelNumber = ModelString[0];
						System.out.println("Model Number from summary: "+ModelNumber);
						TCStatus = eElement.getAttribute("Status");
						if(TCStatus.contentEquals("1")||TCStatus.contentEquals("2")) {
							//TestcaseNameArrry.add(eElement.getAttribute("TCName"));
						}
						//array_tcStatus.add(TCStatus);
						array_tcName.add(ModelNumber);
					}
				}
				System.out.println("array_tcName "+array_tcName);
				int listLength=array_tcName.size();

				for(int i=0;i<listLength;i++)
				{
					HashMapTcStatus.put(array_tcName.get(i), array_tcStatus.get(i));

				}
				for(Map.Entry m: HashMapTcStatus.entrySet())     
				{  
					System.out.println(m.getKey()+" "+m.getValue());   
				}
			}
			else {
				HashMapTcStatus = null;
				Date = null;
				SuiteName = null;
				System.out.println("SummaryReport File not Fond");   
			}
		} 
		catch (Exception e)
		{
			System.out.println(e);
		}
	}

	public static void reportLinkGenration() throws EncryptedDocumentException, IOException
	{
		ArrayList<String> reportLink= new ArrayList<String>();
		String mostRecentITR_Folder=Integer.toString(latestCreatedFolder);
		String itrFolderVal="ITR_"+mostRecentITR_Folder;
		String reportLinPath=dir+"\\"+itrFolderVal+"\\SummaryReport.html";
		System.out.println("array_tcNamefrominfo "+array_tcNamefrominfo);
		System.out.println("array_tcNamefrominfo.size() "+array_tcNamefrominfo.size());
		for(int i=0;i<array_tcNamefrominfo.size();i++)
		{
			reportLink.add(reportLinPath);
			HashMapreportLink.put(array_tcNamefrominfo.get(i), reportLink.get(i));
		}  
		for(excelPrintIterator = 1;excelPrintIterator < 3;excelPrintIterator++) {
			if(excelPrintIterator == 1) {
				printinExcel(reportPathOnLocal);
				System.out.println("Excel written successfully..on Local");
			}
			else {
				printinExcel(reportPath);
				System.out.println("Excel written successfully..on Drive");
			}
		}
	}

	public static void printinExcel(String reportPathOnLocal) throws EncryptedDocumentException, IOException 

	{
		class local
		{	
			
			public void excelPrinting() throws IOException{
				String excelPath; 
				//int FailedCount;
				excelPath = reportPathOnLocal+"_"+Date+".xls";

				File f = new File(excelPath);
				if (f.exists())                       //if Excel file exist 
				{
					FileInputStream inputStream = new FileInputStream(new File(excelPath));

					HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
					HSSFSheet sheet;
/*
					String sheetName;
					int sheetCount = workbook.getNumberOfSheets();
					if(sheetCount != 4) {
						sheetName = workbook.getSheetName(0);
						if(!sheetName.equals("EMEA") ) {
							sheet = workbook.createSheet("EMEA");
						}
						if(!sheetName.equals("Q3_NAM") ) {
							sheet = workbook.createSheet("Q3_NAM");
						}
						if(!sheetName.equals("EU7") ) {
							sheet = workbook.createSheet("EU7");
						}
						if(!sheetName.equals("NAM") ) {
							sheet = workbook.createSheet("NAM");
						}
					}

					if(SuiteName.contains("EMEA") ) {
						sheet = workbook.getSheet("EMEA"); 
					}
					else if(SuiteName.contains("Q3_NAM")) {
						sheet = workbook.getSheet("Q3_NAM"); 
					}
					else if(SuiteName.contains("EU7")) {
						sheet = workbook.getSheet("EU7"); 
					}
					else {
						sheet = workbook.getSheet("NAM"); 
					}
*/
					sheet = workbook.getSheetAt(0);
					
					CellStyle style4_columnHeading = workbook.createCellStyle();
					HSSFFont font = workbook.createFont(); 
					HSSFFont font1 = workbook.createFont();
					HSSFFont font2 = workbook.createFont(); 
					CellStyle style = workbook.createCellStyle();
					CellStyle style1 = workbook.createCellStyle();
					CellStyle style2 = workbook.createCellStyle();
					CellStyle style3 = workbook.createCellStyle();
					CellStyle style4 = workbook.createCellStyle();


					StyleFormating(font , font1 ,font2, style, style1, style2, style3, style4 , style4_columnHeading);

					try { 
						int rowCount = sheet.getLastRowNum();
						if(rowCount < 0)            //if Sheet is empty
						{
							headerPrinting(sheet,style4_columnHeading);
							int rownum;
							rownum=0;
							analysisPrinting(rownum,workbook,sheet,font , font1 ,font2, style, style1, style2, style3, style4);
						}	

						else                      //if Sheet is not empty
						{
							analysisPrinting(rowCount,workbook,sheet,font , font1 ,font2, style, style1, style2, style3, style4);
						}

					}
					catch (Exception e) 
					{
						// TODO: handle exception
						e.printStackTrace();
					}

					inputStream.close();
					System.out.println(excelPath);

					FileOutputStream out= new FileOutputStream(new File(excelPath));

					workbook.write(out);
					workbook.close();
					out.close();
				}
				else								 //if Excel file does not exist 
				{
					HSSFWorkbook workbook = new HSSFWorkbook();
					HSSFSheet sheet; 
					CellStyle style4_columnHeading = workbook.createCellStyle();
					HSSFFont font = workbook.createFont(); 
					HSSFFont font1 = workbook.createFont();
					HSSFFont font2 = workbook.createFont(); 
					CellStyle style = workbook.createCellStyle();
					CellStyle style1 = workbook.createCellStyle();
					CellStyle style2 = workbook.createCellStyle();
					CellStyle style3 = workbook.createCellStyle();
					CellStyle style4 = workbook.createCellStyle();

					sheet = workbook.createSheet("All Model Result List"); 
//					if(SuiteName.contains("EMEA")) {
//						sheet = workbook.createSheet("EMEA"); 
//					}
//					else if(SuiteName.contains("Q3_NAM")) {
//						sheet = workbook.createSheet("Q3_NAM"); 
//					}
//					else if(SuiteName.contains("EU7")) {
//						sheet = workbook.createSheet("EU7"); 
//					}
//					else {
//						sheet = workbook.createSheet("NAM"); 
//					}

					StyleFormating(font , font1 ,font2, style, style1, style2, style3,style4, style4_columnHeading);	
					headerPrinting(sheet,style4_columnHeading);	
					try {
						int rownum;
						rownum=0;
						analysisPrinting(rownum,workbook,sheet,font , font1 ,font2, style, style1, style2, style3, style4);
					}
					catch (Exception e) 
					{
						// TODO: handle exception
						e.printStackTrace();
					}

					System.out.println(excelPath);

					FileOutputStream out= new FileOutputStream(new File(excelPath));

					workbook.write(out);
					workbook.close();
					out.close();
				}
			}

			public void StyleFormating(HSSFFont font ,HSSFFont font1 ,HSSFFont font2, CellStyle style, CellStyle style1, CellStyle style2, CellStyle style3, CellStyle style4 , CellStyle style4_columnHeading) {

				style.setFont(font);
				style1.setFont(font);
				style2.setFont(font);
				style3.setFont(font);
				style4.setFont(font);

				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
				style.setBorderTop(BorderStyle.THIN);
				style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				style.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style.setTopBorderColor(IndexedColors.BLACK.getIndex());

				style4_columnHeading.setBorderBottom(BorderStyle.THIN);
				style4_columnHeading.setBorderLeft(BorderStyle.THIN);
				style4_columnHeading.setBorderRight(BorderStyle.THIN);
				style4_columnHeading.setBorderTop(BorderStyle.THIN);
				style4_columnHeading.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				style4_columnHeading.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				style4_columnHeading.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style4_columnHeading.setTopBorderColor(IndexedColors.BLACK.getIndex());

				style2.setBorderBottom(BorderStyle.THIN);
				style2.setBorderLeft(BorderStyle.THIN);
				style2.setBorderRight(BorderStyle.THIN);
				style2.setBorderTop(BorderStyle.THIN);
				style2.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				style2.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				style2.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style2.setTopBorderColor(IndexedColors.BLACK.getIndex());

				style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
				style1.setFillPattern(FillPatternType.SOLID_FOREGROUND); 
				style4_columnHeading.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
				style4_columnHeading.setFillPattern(FillPatternType.SOLID_FOREGROUND); 
				style3.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				
				style4.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);

				font.setFontName("Calibri");
				font.setFontHeightInPoints((short)12);
				font2.setFontName("Calibri");
				font2.setFontHeightInPoints((short)14);
				font2.setBold(true);
				style4_columnHeading.setFont(font2);
			}

			public void headerPrinting(HSSFSheet sheet, CellStyle style4_columnHeading) {

				int rownum0=0;
				int cellnum0 = 0;
				Row row0 = sheet.createRow(rownum0);
				if(excelPrintIterator == 2) {
					Cell cellsRnO = row0.createCell(cellnum0);
					cellnum0++;
					Cell cellModalities = row0.createCell(cellnum0);
					cellnum0++;
					Cell cellMarkets = row0.createCell(cellnum0);
					cellnum0++;
					Cell cellCountry = row0.createCell(cellnum0);
					cellnum0++;
					cellsRnO.setCellValue("Sr No.");
					cellsRnO.setCellStyle(style4_columnHeading);
					cellModalities.setCellValue("Modalities");
					cellModalities.setCellStyle(style4_columnHeading);
					cellMarkets.setCellValue("Market");
					cellMarkets.setCellStyle(style4_columnHeading);
					cellCountry.setCellValue("Country");
					cellCountry.setCellStyle(style4_columnHeading);
				}
				Cell cell00 = row0.createCell(cellnum0);
				if(excelPrintIterator == 2) {
					cellnum0++;
					Cell ImpactedModels = row0.createCell(cellnum0);				
					ImpactedModels.setCellValue("Impacted Models");	
					ImpactedModels.setCellStyle(style4_columnHeading);
				}
				cellnum0++;
				Cell cell01 = row0.createCell(cellnum0);
				cellnum0++;
				Cell cell02 = row0.createCell(cellnum0);
				cellnum0++;
				Cell cell03 = row0.createCell(cellnum0);
				cellnum0++;
				Cell SequencedFailureAnalysis = row0.createCell(cellnum0);
				if(excelPrintIterator == 1) {	
				cellnum0++;
				Cell cell04 = row0.createCell(cellnum0);
				cellnum0++;
				Cell cell05 = row0.createCell(cellnum0);
				cellnum0++;
				Cell cell06 = row0.createCell(cellnum0);
				cellnum0++;
				Cell cell07 = row0.createCell(cellnum0);
				cell04.setCellValue("Extra Error on UI");
				cell05.setCellValue("Expected Error Not on UI");
				cell06.setCellValue("Errors Not Eradicated");
				cell07.setCellValue("Errors even after selecting");
				cell04.setCellStyle(style4_columnHeading);
				cell05.setCellStyle(style4_columnHeading);
				cell06.setCellStyle(style4_columnHeading);
				cell07.setCellStyle(style4_columnHeading);
				}
				
				if(excelPrintIterator == 2) {	
				cellnum0++;
				Cell SuiteName = row0.createCell(cellnum0);				
				SuiteName.setCellValue("Suite Name");					
				SuiteName.setCellStyle(style4_columnHeading);
				cellnum0++;
				Cell FailedCount = row0.createCell(cellnum0);				
				FailedCount.setCellValue("Failed Count");					
				FailedCount.setCellStyle(style4_columnHeading);
//				cellnum0++;
//				Cell FailedCount = row0.createCell(cellnum0);				
//				FailedCount.setCellValue("Failed Count");					
//				FailedCount.setCellStyle(style4_columnHeading);
				}
				if(excelPrintIterator == 1) {
					cellnum0++;
					Cell cell08 = row0.createCell(cellnum0);
					cellnum0++;
					Cell cell09 = row0.createCell(cellnum0);
					cell08.setCellValue("Error Screenshots");
					cell09.setCellValue("Report Link");
					cell08.setCellStyle(style4_columnHeading);
					cell09.setCellStyle(style4_columnHeading);
//					cellnum0++;
//					Cell FailedCountLocal = row0.createCell(cellnum0);
//					FailedCountLocal.setCellValue("Failed Count");
//					FailedCountLocal.setCellStyle(style4_columnHeading);
				}
				//System.out.println("cellnum0 "+cellnum0);
				cell00.setCellValue("Model Number");
				cell01.setCellValue("QuoteNumber_"+Date);
				cell02.setCellValue("Status_"+Date);
				cell03.setCellValue("Failure Analysis_"+Date);
				SequencedFailureAnalysis.setCellValue("Sequenced Failure Analysis"+Date);
				

				
				cell00.setCellStyle(style4_columnHeading);
				cell01.setCellStyle(style4_columnHeading);
				cell02.setCellStyle(style4_columnHeading);
				cell03.setCellStyle(style4_columnHeading);
				SequencedFailureAnalysis.setCellStyle(style4_columnHeading);

				
			}

			public void analysisPrinting(int rownum, HSSFWorkbook workbook,HSSFSheet sheet, HSSFFont font ,HSSFFont font1 ,HSSFFont font2, CellStyle style, CellStyle style1, CellStyle style2, CellStyle style3, CellStyle style4) throws IOException {

				System.out.println("analysisPrinting func");
				int cellnum = 0;
				int flag;
				if(excelPrintIterator == 1)
					flag = 0;
				else
					flag = 1;
				int ab = 0;	

				Iterator <String> it = TCexecutionStatus.keySet().iterator();
				LastSrNoCount = sheet.getLastRowNum();
				int FailedCountPrinting = 1;
			
				
				
				
				while(it.hasNext())  
				{   
					Row row = sheet.createRow(++rownum);
					if(FailedCountPrinting == 1) {
						FailedCountPrinting++;
						if(excelPrintIterator == 2) {
							Row row1 = sheet.getRow(1); 
							Cell previousFailedCountCell = row1.getCell(11);
							if(previousFailedCountCell == null) {
								previousFailedCountCell = row1.createCell(11);
								previousFailedCountCell.setCellValue(0);
								previousFailedCountCell = row1.getCell(11);
							}
							System.out.println("previousFailedCountCellLocal "+previousFailedCountCell);
							Double previousFailedCount = previousFailedCountCell.getNumericCellValue();
							//previousFailedCountCell.get
							System.out.println("previousFailedCount "+previousFailedCount);
							int previousFailedCountInt = (int)Math.round(previousFailedCount);
							System.out.println("previousFailedCountInt "+previousFailedCountInt);
							previousFailedCountInt = previousFailedCountInt+FailedCount;
							System.out.println("previousFailedCountInt "+previousFailedCountInt);
							previousFailedCountCell.setCellValue(previousFailedCountInt);
							previousFailedCountCell.setCellStyle(style);
							}
//						else {
//							Row row1 = sheet.getRow(1); 
//							Cell previousFailedCountCell = row1.getCell(13);
//							System.out.println("previousFailedCountCell "+previousFailedCountCell);
//							Double previousFailedCount = previousFailedCountCell.getNumericCellValue();
//							System.out.println("previousFailedCount "+previousFailedCount);
//							int previousFailedCountInt = (int)Math.round(previousFailedCount);
//							System.out.println("previousFailedCountInt "+previousFailedCountInt);
//							previousFailedCountCell.setCellValue(previousFailedCountInt+FailedCount);
//							previousFailedCountCell.setCellStyle(style);
//						}
					}
					String key= it.next();
					String value  = HashMapQuoteCapt.get(key);
					String value2 = TCexecutionStatus.get(key);
					//String value2 = HashMapTcStatus.get(key);
					String value4 = HashMapreportLink.get(key);
					String value3 = HashMapErrorMessages.get(key);
					String SequencedErrors = SeceuencedFailureMap.get(key);
					String value5 = HashMapScreenshot.get(key);
					String value6 = ExtraRules.get(key);
					String value7 = ExpectedRules.get(key);
					String value8 = NotEradicatedRules.get(key);
					String value9 = evenAfterRules.get(key);
					String valueCountryCode = CountryCodeMAp.get(key);

					System.out.println("key :"+key);
					final Hyperlink href = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
					href.setAddress(value4);
					if(excelPrintIterator == 2) 
					{
						Cell cellSrNo = row.createCell(cellnum);
						cellnum++;
						Cell cellModalities = row.createCell(cellnum);
						cellnum++;
						Cell cellMarkets = row.createCell(cellnum);
						cellnum++;
						Cell cellCountry = row.createCell(cellnum);
						cellnum++;
						cellSrNo.setCellValue(++LastSrNoCount);
						cellSrNo.setCellStyle(style);
						cellModalities.setCellValue(ModalityExtraction(key));
						cellModalities.setCellStyle(style);
						cellMarkets.setCellValue(MarketExtraction(valueCountryCode));
						cellMarkets.setCellStyle(style);
						cellCountry.setCellValue(valueCountryCode);
						cellCountry.setCellStyle(style);
					}
					Cell cell = row.createCell(cellnum);
					if(excelPrintIterator == 2) {
						cellnum++;
						Cell ImpactedModels = row.createCell(cellnum);
						if(ImpactedModelPresent(key)) {
							ImpactedModels.setCellValue("Yes");
							ImpactedModels.setCellStyle(style);
						}
						else {
							ImpactedModels.setCellValue(".");
							ImpactedModels.setCellStyle(style);
						}
					}
					cellnum++;
					Cell cell2 = row.createCell(cellnum);
					cellnum++;
					Cell cell3 = row.createCell(cellnum);
					cellnum++;
					Cell cell4 = row.createCell(cellnum);
					cellnum++;
					Cell SequencedErrorsCell = row.createCell(cellnum);
				//	System.out.println("SequencedErrors "+SequencedErrors);

					if(excelPrintIterator == 1) 
					{
					cellnum++;
					Cell cell5 = row.createCell(cellnum);
					cellnum++;
					Cell cell6 = row.createCell(cellnum);
					cellnum++;
					Cell cell7 = row.createCell(cellnum);
					cellnum++;
					Cell cell8 = row.createCell(cellnum);
					cell5.setCellValue(value6);
					cell6.setCellValue(value7);
					cell7.setCellValue(value8);
					cell8.setCellValue(value9);
					cell5.setCellStyle(style);
					cell6.setCellStyle(style);
					cell7.setCellStyle(style);
					cell8.setCellStyle(style);
					}

					if(excelPrintIterator == 1) {
						cellnum++;
						Cell cell9 = row.createCell(cellnum);

						if(value5 == "---") {
							cell9.setCellStyle(style);
							cell9.setCellValue("-");
						}
						else if(value5 == "-/-")
						{
							cell9.setCellStyle(style);
							cell9.setCellValue("No screenshots Taken");
						}
						else {
							cell9.setCellStyle(style2);
							final Hyperlink href1 = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
							href1.setAddress(value5);
							font1.setColor(IndexedColors.BLUE.getIndex());
							style2.setFont(font1);
							String Screenshot="Click_Here";
							cell9.setCellValue(Screenshot);
							cell9.setHyperlink((org.apache.poi.ss.usermodel.Hyperlink) href1);
						} 
						if(ab == 0) {
							cellnum++;
							Cell cell10 = row.createCell(cellnum);
							font1.setColor(IndexedColors.BLUE.getIndex());
							style2.setFont(font1);
							String reportPath="Click_Here";
							cell10.setCellValue(reportPath);
							cell10.setHyperlink((org.apache.poi.ss.usermodel.Hyperlink) href);
							cell10.setCellStyle(style2);
						}
						ab++;
					}
					String keyaray[] = key.split("-");
					//cell.setCellValue(keyaray[0]);
					cell.setCellValue(key);
					cell2.setCellValue(value);
					System.out.println("value2 "+value2);
					
			
					
					if(value2.equals("0"))
					{
						String status="Pass";
						cell3.setCellValue(status);
						cell3.setCellStyle(style);
					}
					else if(value2.equals("1"))
					{
						
						String status="Fail";
						cell3.setCellValue(status);
						if(excelPrintIterator == 1)
							cell3.setCellStyle(style3);
						else
							cell3.setCellStyle(style);
					}

//					else
//					{
//						String status="Fail";
//						cell3.setCellValue(status);
//						if(excelPrintIterator == 1)
//							cell3.setCellStyle(style3);
//						else
//							cell3.setCellStyle(style);
//					}

					ArrayList<String> AllTypeRules = new ArrayList<String>();
					AllTypeRules.add(value6);
					AllTypeRules.add(value7);
					AllTypeRules.add(value8);
					AllTypeRules.add(value9);
					//ArrayList<String> AllRules = new ArrayList<String>();
					String AllRules ="";
					int ruleCount = 0;
					boolean RulesErrorExist = false;
					System.out.println("AllTypeRules"+AllTypeRules);
					for(String i: AllTypeRules) {
						ruleCount++;
						if(i.contentEquals("-")||i.contentEquals("[]"))
						{
							continue;			        		
						}
						else {
							RulesErrorExist = true;

							if( ruleCount == 1) {
								AllRules = "\n\nExtra Rules on UI:\n"+i;
							}
							if( ruleCount == 2) {
								AllRules = AllRules+"\n\nExpected Rules not on ui:\n"+i;
							}
							if( ruleCount == 3) {
								AllRules = AllRules+"\n\nRules not Eradicated:\n"+i;
							}
							if( ruleCount == 4) {
								AllRules = AllRules+"\n\nEven after selecting all options Rules exist:\n"+i;
							}
						}
					}
					if(value3.contains("Wrenchicon Shows pending")) {
						value3 = value3.replace("Wrenchicon Shows pending configuration.", "");
						if(RulesErrorExist == true) {
							cell4.setCellValue(value3+AllRules+"\n\nWrenchicon Shows pending configuration. Hence, Quote not configured successfully.");
							System.out.println("RulesErrorExist");
							cell4.setCellStyle(style);
						}
						else {
							cell4.setCellValue(value3+"\nWrenchicon Shows pending configuration. Hence, Quote not configured successfully.");
							System.out.println("RulesErrornotExist");
							cell4.setCellStyle(style);
						}
					}
					else if(value3.contains("Go - to - Pricing Disabled")||value3.contains("not found on Catlog page")) {
						if(RulesErrorExist == true) {
							cell4.setCellValue(value3+AllRules+"\n\nHence, Quote not configured successfully.");
							System.out.println("RulesErrorExist");
							cell4.setCellStyle(style);
						}
						else {
							cell4.setCellValue(value3+"\n\nHence, Quote not configured successfully.");
							System.out.println("RulesErrornotExist");
							cell4.setCellStyle(style);
						}
					}
					else if(value3.contentEquals("No error")) {
						if(RulesErrorExist == true) {
							cell4.setCellValue("Failed:"+AllRules+"\n\nBut Quote configured successfully.");
							System.out.println("RulesErrorExist");
							cell4.setCellStyle(style);
						}
						else {
							cell4.setCellValue("-");
							System.out.println("RulesErrornotExist");
							cell4.setCellStyle(style);
						}
					}
					else if(value3.contentEquals("Failed: ")) {
						if(RulesErrorExist == true) {
							cell4.setCellValue(value3+AllRules+"\n\nBut Quote configured successfully.");
							System.out.println("RulesErrorExist");
							cell4.setCellStyle(style);
						}
						else {
							System.out.println("Failed with some other reason");
							cell4.setCellValue("Failed with some other reason.\nBut Quote configured successfully");
							System.out.println("RulesErrornotExist");
							cell4.setCellStyle(style4);
							for(String i: TestcaseNameArrry) {
								String[] ModelNUMString = i.split("_", 0);
								String ModelNumber = ModelNUMString[0];
								System.out.println("Model Number from summary: "+ModelNumber);
								if(ModelNumber.contentEquals(key)) {
									FailedTestcaseNameArrry.add(i);
								}
							}
						}
					}
					else {
						if(RulesErrorExist == true) {
							cell4.setCellValue(value3+AllRules+"\n\nBut Quote configured successfully.");
							System.out.println("RulesErrorExist");
							cell4.setCellStyle(style);
						}
						else {
							cell4.setCellValue(value3+"\n\nBut Quote configured successfully.");
							System.out.println("RulesErrornotExist");
							cell4.setCellStyle(style);
						}
					}
//--------------------------------------------------------------------------------------------------------------------------------
					if(SequencedErrors.contains("Wrenchicon Shows pending")) {
						SequencedErrors = SequencedErrors.replace("Wrenchicon Shows pending configuration.", "");
							SequencedErrorsCell.setCellValue(SequencedErrors+"\nWrenchicon Shows pending configuration. Hence, Quote not configured successfully.");
							SequencedErrorsCell.setCellStyle(style);
					}			
					else if(SequencedErrors.contains("Go - to - Pricing Disabled")||SequencedErrors.contains("not found on Catlog page")) {
						SequencedErrorsCell.setCellValue(SequencedErrors+"\nHence, Quote not configured successfully.");
						SequencedErrorsCell.setCellStyle(style);
					}				
					else if(SequencedErrors.contentEquals("Failed: ")) {
							System.out.println("Failed with some other reason");
							SequencedErrorsCell.setCellValue("Failed with some other reason.\nBut Quote configured successfully");
							SequencedErrorsCell.setCellStyle(style4);
//							for(String i: TestcaseNameArrry) {
//								String[] ModelNUMString = i.split("_", 0);
//								String ModelNumber = ModelNUMString[0];
//								System.out.println("Model Number from summary: "+ModelNumber);
//								if(ModelNumber.contentEquals(key)) {
//									FailedTestcaseNameArrry.add(i);
//								}
//							}
					}
					else if(SequencedErrors.contentEquals("-")) {
						SequencedErrorsCell.setCellValue(SequencedErrors);
						SequencedErrorsCell.setCellStyle(style);
					}
					else {
						SequencedErrorsCell.setCellValue(SequencedErrors+"\nBut Quote configured successfully.");
						SequencedErrorsCell.setCellStyle(style);
					}


					Cell cell11 = null;
					if(flag==1)
					{	
						cellnum++;
						cell11 = row.createCell(cellnum);
						cell11.setCellValue(SuiteName);
						cell11.setCellStyle(style);

					}
					if(flag==2) {
						cellnum++;
						cell11 = row.createCell(cellnum);
						cell11.setCellValue("Date:"+Date+" Time:"+ExecutionTime);
						cell11.setCellStyle(style);
					}
					flag++;
					cellnum=0;

					cell.setCellStyle(style);
					cell2.setCellStyle(style);
					//	cell4.setCellStyle(style);

					
				

				}
				if(excelPrintIterator == 1) {		//on LOCAL
					sheet.autoSizeColumn(0);
					sheet.setColumnWidth(1, 5000);
					sheet.setColumnWidth(2,5000);
					sheet.setColumnWidth(3, 5000);
					sheet.setColumnWidth(4, 4500);
					sheet.setColumnWidth(5, 4500);
					sheet.setColumnWidth(6, 4500);
					sheet.setColumnWidth(7, 6000);
					sheet.setColumnWidth(8, 6000);
					sheet.setColumnWidth(9, 6000);
					sheet.setColumnWidth(10, 7500);
					//sheet.setColumnWidth(10, 2500);
					//sheet.setColumnWidth(11, 7500);
					//sheet.setColumnWidth(12, 7500);
				}
				if(excelPrintIterator == 2) {		//on DRIVE
					sheet.setColumnWidth(0, 2000);
					sheet.autoSizeColumn(1);
					sheet.autoSizeColumn(2);
					sheet.autoSizeColumn(3);
					sheet.autoSizeColumn(4);
					sheet.setColumnWidth(5, 6000);
					sheet.setColumnWidth(6, 4500);
					sheet.setColumnWidth(7, 6000);
					sheet.setColumnWidth(8, 6000);
					sheet.setColumnWidth(9, 7500);
					sheet.setColumnWidth(10, 6000);
					sheet.setColumnWidth(11, 2500);
				}

			}

			public String ModalityExtraction(String ModelNumberKey) throws IOException {
				
				String keyaray[] = ModelNumberKey.split("-");
				ModelNumberKey = keyaray[0];
			
				String Modalitysheet1;
				Modalitysheet1 = Modalitysheet+".xls";
				File f = new File(Modalitysheet1);
				String Modalitystr = "NA";
				if (f.exists())                       //if Excel file exist 
				{
					FileInputStream inputStream = new FileInputStream(new File(Modalitysheet1));
					HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
					HSSFSheet sheet = workbook.getSheet("Modalities");  
					int rowCount = sheet.getLastRowNum();
					for(i=1;i<=rowCount;i++) {
						Row row = sheet.getRow(i); 
						String ModNumberstr= row.getCell(0).toString();  
						if(ModNumberstr.contains(".0")) {
							ModNumberstr = ModNumberstr.replace(".0", "");
						}
						if(ModNumberstr.contains(ModelNumberKey)) {
							Cell Modality = row.getCell(1);
							System.out.println("Modality "+Modality);
							Modalitystr = Modality.toString();	
							break;
						}

					}
					workbook.close();
					inputStream.close();
				}
				else {
					System.out.println("Modality File Not Fouund");
				}
				System.out.println("Modality for "+ModelNumberKey +": "+Modalitystr);		
				return Modalitystr;
			}

			public String MarketExtraction(String CountryCodeKey) throws IOException {
				String Modalitysheet1;
				Modalitysheet1 = Modalitysheet+".xls";
				File f = new File(Modalitysheet1);
				String Marketstr = "NA";
				if (f.exists())                       //if Excel file exist 
				{
					FileInputStream inputStream = new FileInputStream(new File(Modalitysheet1));
					HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
					HSSFSheet sheet = workbook.getSheet("Market");  
					int rowCount = sheet.getLastRowNum();
					for(i=1;i<=rowCount;i++) {
						Row row = sheet.getRow(i); 
						String CountryCodestr= row.getCell(0).toString();  
						if(CountryCodestr.contains(".0")) {
							CountryCodestr = CountryCodestr.replace(".0", "");
						}
						if(CountryCodestr.contains(CountryCodeKey)) {
							Cell Market = row.getCell(1);
							System.out.println("Market "+Market);
							Marketstr = Market.toString();	
							break;
						}

					}
					workbook.close();
					inputStream.close();
				}
				else {
					System.out.println("Modality File Not Fouund");
				}
				System.out.println("Market for "+CountryCodeKey+" : "+Marketstr);		
				return Marketstr;
			}
			
			public boolean ImpactedModelPresent(String ModelNumber) throws IOException  
			{
				String keyaray[] = ModelNumber.split("-");
				ModelNumber = keyaray[0];
				boolean ModelPresent = false;
				String Modalitysheet1;
				Modalitysheet1 = Modalitysheet+".xls";
				File f = new File(Modalitysheet1);
				if (f.exists())                       //if Excel file exist 
				{
					FileInputStream inputStream = new FileInputStream(new File(Modalitysheet1));
					HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
					HSSFSheet sheet = workbook.getSheet("ImpactedModels");  
					int rowCount = sheet.getLastRowNum();
					for(i=0;i<=rowCount;i++) 
					{
						Row row = sheet.getRow(i); 
						String ImpactedModel= row.getCell(0).toString();  
						if(ImpactedModel.contains(".0"))
						{
							ImpactedModel = ImpactedModel.replace(".0", "");
						}
						if(ModelNumber.contentEquals(ImpactedModel)) 
						{
							ModelPresent = true;
							break;
						}
						else
							ModelPresent = false;

					}
					workbook.close();
					inputStream.close();
				}
				else {
					System.out.println("Modality File Not Fouund");
				}
				return ModelPresent;
			}
			
		}
		new local().excelPrinting();
	}

}