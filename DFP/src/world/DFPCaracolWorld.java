package world;

import java.awt.TextArea;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.Properties;

import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.api.ads.common.lib.auth.OfflineCredentials;
import com.google.api.ads.common.lib.auth.OfflineCredentials.Api;
import com.google.api.ads.dfp.axis.factory.DfpServices;
import com.google.api.ads.dfp.axis.utils.v201605.DateTimes;
import com.google.api.ads.dfp.axis.utils.v201605.ReportDownloader;
import com.google.api.ads.dfp.axis.utils.v201605.StatementBuilder;
import com.google.api.ads.dfp.axis.v201605.Column;
import com.google.api.ads.dfp.axis.v201605.Date;
import com.google.api.ads.dfp.axis.v201605.DateRangeType;
import com.google.api.ads.dfp.axis.v201605.Dimension;
import com.google.api.ads.dfp.axis.v201605.DimensionAttribute;
import com.google.api.ads.dfp.axis.v201605.ExportFormat;
import com.google.api.ads.dfp.axis.v201605.Order;
import com.google.api.ads.dfp.axis.v201605.OrderPage;
import com.google.api.ads.dfp.axis.v201605.OrderServiceInterface;
import com.google.api.ads.dfp.axis.v201605.ReportDownloadOptions;
import com.google.api.ads.dfp.axis.v201605.ReportJob;
import com.google.api.ads.dfp.axis.v201605.ReportQuery;
import com.google.api.ads.dfp.axis.v201605.ReportServiceInterface;
import com.google.api.ads.dfp.lib.client.DfpSession;
import com.google.api.client.auth.oauth2.Credential;
import com.google.common.io.Files;
import com.google.common.io.Resources;

import gui.GUI;

public class DFPCaracolWorld {
	
	
	
	
	
	private ArrayList orderId;
	public static String PROPERTIES = "config.properties";
	public static String parentPath = "";
	public static String reportPath = "";
	public static String plantPath = "";
	public TextArea textArea;
	private GUI gui;
	
	
	public DFPCaracolWorld(DfpServices dfpServices, DfpSession session, TextArea textArea , GUI gui , File p) throws Exception
	{
		
//		redirectSystemStreams();
		this.gui = gui;
		orderId = new ArrayList();
		
		this.textArea = textArea;
		
		readProp(textArea , p);
		
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		
		 // Get the ReportService.
	    ReportServiceInterface reportService = dfpServices.get(session, ReportServiceInterface.class);
	    
	    OrderServiceInterface orderService = dfpServices.get(session, OrderServiceInterface.class);
	    

	    // Create a statement to select all orders.
	    StatementBuilder statementBuilder = new StatementBuilder().where("NOT NAME like '%Autopauta%' and STATUS = 'Completed'")
	        .orderBy("id ASC")
	        .limit(StatementBuilder.SUGGESTED_PAGE_LIMIT);

	    // Default for total result set size.
	    int totalResultSetSize = 0;

	  
	      // Get orders by statement.
	      OrderPage page =  orderService.getOrdersByStatement(statementBuilder.toStatement());
int h  =0;
	      if (page.getResults() != null) {
	        totalResultSetSize = page.getTotalResultSetSize();
	        int i = page.getStartIndex();
	        for (Order order : page.getResults()) {
	        	
	        	//yyyy/MM/dd
	        	if(order.getEndDateTime()!= null)
	        	{
	        		
	        	
	        	String dat = order.getEndDateTime().getDate().getYear()+"/"+ order.getEndDateTime().getDate().getMonth() +"/"+ order.getEndDateTime().getDate().getDay();
	        	if(dat.equals(getYesterdayDateString()))
	        			{
	        		orderId.add(order);
	        		gui.pintarTextArea((i++)+")Order with ID"+order.getId()+" And name "+ order.getName());
	        		h++;
	        			}
	        	
	        	}
	        	
	         
	        }
	      }

	      statementBuilder.increaseOffsetBy(StatementBuilder.SUGGESTED_PAGE_LIMIT);
	     while (statementBuilder.getOffset() < totalResultSetSize);

	     gui.pintarTextArea("Number of results found:"+totalResultSetSize );
	     gui.pintarTextArea("Number Of results processed: "+h);
	     
	  

//	    // Create report query.
	    ReportQuery reportQueryReach = new ReportQuery();
	    ReportQuery reportQueryGeneral = new ReportQuery();
	    ReportQuery reportQueryGeo = new ReportQuery();
	    ReportQuery reportQueryCreative = new ReportQuery();
	    ReportQuery reportQueryDevices = new ReportQuery();
	    orderId = deduplicateArray();
	    for(int i = 0; i < orderId.size(); i++)
	    {
	    	ArrayList productsByActualOrder = new ArrayList<Product>();
	    	
//	    	113957316 

	    	Order orden = (Order) orderId.get(i);
	    	StatementBuilder statement = new StatementBuilder()    			
			        .where("ORDER_ID ="+orden.getId());
	    	gui.pintarTextArea("Retrieving data for:"+orden.getName());
	    	gui.pintarTextArea("report "+i+ " of " + orderId.size());
	    	
		    reportQueryGeneral.setStatement(statement.toStatement());
		    reportQueryDevices.setStatement(statement.toStatement());
		    reportQueryReach.setStatement(statement.toStatement());
		    reportQueryGeo.setStatement(statement.toStatement());
		    reportQueryCreative.setStatement(statement.toStatement());
		    
		    
		    reportQueryReach.setDimensions(new Dimension[] {Dimension.ORDER_ID, Dimension.ORDER_NAME,Dimension.LINE_ITEM_ID, Dimension.LINE_ITEM_NAME});
		    reportQueryReach.setColumns(new Column[] {Column.REACH, Column.REACH_FREQUENCY});
		    
		    	    
		    reportQueryGeneral.setDimensions(new Dimension[] {Dimension.DATE,Dimension.ADVERTISER_ID,Dimension.ADVERTISER_NAME , Dimension.ORDER_ID, Dimension.ORDER_NAME , Dimension.AD_UNIT_ID, Dimension.AD_UNIT_NAME, Dimension.LINE_ITEM_ID, Dimension.LINE_ITEM_NAME	});
		    reportQueryGeneral.setColumns(new Column[] {  Column.AD_SERVER_IMPRESSIONS, Column.TOTAL_LINE_ITEM_LEVEL_CLICKS});
		    		    
		    reportQueryGeneral.setDimensionAttributes(new DimensionAttribute[] {DimensionAttribute.ORDER_TRAFFICKER, DimensionAttribute.ORDER_START_DATE_TIME,
		            DimensionAttribute.ORDER_END_DATE_TIME, DimensionAttribute.ORDER_SALESPERSON,DimensionAttribute.ORDER_BOOKED_CPC, DimensionAttribute.ORDER_BOOKED_CPM
		    		,});
		    
		    reportQueryGeo.setDimensions(new Dimension[] { Dimension.CITY_NAME,Dimension.ORDER_ID, Dimension.ORDER_NAME, Dimension.LINE_ITEM_ID, Dimension.LINE_ITEM_NAME});
		    reportQueryGeo.setColumns(new Column[] {Column.AD_SERVER_IMPRESSIONS, Column.TOTAL_LINE_ITEM_LEVEL_CLICKS});
			  
		
		    reportQueryCreative.setDimensions(new Dimension[] { Dimension.CREATIVE_ID, Dimension.CREATIVE_NAME, Dimension.CREATIVE_SIZE});
		    reportQueryCreative.setColumns(new Column[] {Column.AD_SERVER_IMPRESSIONS, Column.AD_SERVER_CLICKS});
		    
		  
		    reportQueryDevices.setDimensions(new Dimension[] {Dimension.DATE, Dimension.DEVICE_CATEGORY_ID,Dimension.DEVICE_CATEGORY_NAME	});
		    reportQueryDevices.setColumns(new Column[] {  Column.AD_SERVER_IMPRESSIONS, Column.TOTAL_LINE_ITEM_LEVEL_CLICKS});
		    		    
		  
		    // Set the dynamic date range type or a custom start and end date that is
		    // the beginning of the week (Sunday) to the end of the week (Saturday), or
		    // the first of the month to the end of the month.
		    reportQueryReach.setDateRangeType(DateRangeType.REACH_LIFETIME);
		    reportQueryGeneral.setDateRangeType(DateRangeType.CUSTOM_DATE);
		    reportQueryGeo.setDateRangeType(DateRangeType.CUSTOM_DATE);
		    reportQueryCreative.setDateRangeType(DateRangeType.CUSTOM_DATE);
		    reportQueryDevices.setDateRangeType(DateRangeType.CUSTOM_DATE);
		    
		    
String endDateuser = JOptionPane.showInputDialog("Fecha final : p.e.: 2000-05-01 - yyyy-MM-dd");
		    
		    reportQueryCreative.setStartDate(
		    	      DateTimes.toDateTime("2000-05-01T00:00:00", "America/New_York").getDate());
		    	  reportQueryCreative.setEndDate(
		    	      DateTimes.toDateTime(endDateuser+"T00:00:00", "America/New_York").getDate());
		   
		    reportQueryGeo.setStartDate(
		    	      DateTimes.toDateTime("2000-05-01T00:00:00", "America/New_York").getDate());
		    	  reportQueryGeo.setEndDate(
		    	      DateTimes.toDateTime(endDateuser+"T00:00:00", "America/New_York").getDate());
		   
		    reportQueryGeneral.setStartDate(
		    	      DateTimes.toDateTime("2000-05-01T00:00:00", "America/New_York").getDate());
		    	  reportQueryGeneral.setEndDate(
		    	      DateTimes.toDateTime(endDateuser+"T00:00:00", "America/New_York").getDate());
		   
		  	    reportQueryDevices.setStartDate(
			    	      DateTimes.toDateTime("2000-05-01T00:00:00", "America/New_York").getDate());
		  	  reportQueryDevices.setEndDate(
			    	      DateTimes.toDateTime(endDateuser+"T00:00:00", "America/New_York").getDate());
		   
		 
		   
		    // Create report job.
		    ReportJob reportJob = new ReportJob();
		    ReportJob reportJobReach = new ReportJob();
		    ReportJob reportJobGeo = new ReportJob();
		    ReportJob reportJobCreative = new ReportJob();
		    ReportJob reportJobDevices = new ReportJob();
		    reportJob.setReportQuery(reportQueryGeneral);
		    reportJobReach.setReportQuery(reportQueryReach);
		    reportJobGeo.setReportQuery(reportQueryGeo);
		    reportJobCreative.setReportQuery(reportQueryCreative);
		    reportJobDevices.setReportQuery(reportQueryDevices);
		   

		    // Run report job.
		    reportJob = reportService.runReportJob(reportJob);
		    reportJobReach = reportService.runReportJob(reportJobReach);
		    reportJobGeo = reportService.runReportJob(reportJobGeo);
		    reportJobCreative = reportService.runReportJob(reportJobCreative);
		    reportJobDevices = reportService.runReportJob(reportJobDevices);
		 

		    // Create report downloader.
		    ReportDownloader reportDownloader = new ReportDownloader(reportService, reportJob.getId());
		    ReportDownloader reportDownloaderReach = new ReportDownloader(reportService, reportJobReach.getId());
		    ReportDownloader reportDownloaderGeo= new ReportDownloader(reportService, reportJobGeo.getId());
		    ReportDownloader reportDownloaderCreative= new ReportDownloader(reportService, reportJobCreative.getId());
		    ReportDownloader reportDownloaderDevices= new ReportDownloader(reportService, reportJobDevices.getId());
		

		    // Wait for the report to be ready.
		    reportDownloader.waitForReportReady();
		    reportDownloaderReach.waitForReportReady();
		    reportDownloaderGeo.waitForReportReady();
		    reportDownloaderCreative.waitForReportReady();
		    reportDownloaderDevices.waitForReportReady();
		    

		    // Change to your file location.
//		    JFileChooser chooser = new JFileChooser();
//		    chooser.setCurrentDirectory(new java.io.File("."));
//		    String title = "Select destiny path";
//		    File path = null;
//		    
//		    chooser.setDialogTitle(title);
//		    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
//		    if(chooser.showOpenDialog(chooser)== JFileChooser.APPROVE_OPTION)
//		    {
//		    	path = chooser.getSelectedFile();
//		    }
//		   
//		    
		    File fail  = new File(parentPath);
		    
//		    File file = File.createTempFile(orden.getName(), ".csv" , fail);
		    File file = new File(reportPath+orden.getName()+".xlsx");
		    File fileReach = new File(reportPath+orden.getName()+"Reach.xlsx");
		    File fileGeo = new File(reportPath+orden.getName()+"Geo.xlsx");
		    File fileCreative = new File(reportPath+orden.getName()+"Creative.xlsx");
		    File fileDevices= new File(reportPath+orden.getName()+"Devices.xlsx");
		    
		    
		    
		    gui.pintarTextArea("Downloading reports to: " +reportPath+ " File: "+ file.toString());
		    gui.pintarTextArea("Downloading reports to: " +reportPath+ " File: "+ fileReach.toString());
		    gui.pintarTextArea("Downloading reports to: " +reportPath+ " File: "+ fileGeo.toString());
		    gui.pintarTextArea("Downloading reports to: " +reportPath+ " File: "+ fileCreative.toString());
		    gui.pintarTextArea("Downloading reports to: " +reportPath+ " File: "+ fileDevices.toString());
		    

		    // Download the report.
		    ReportDownloadOptions options = new ReportDownloadOptions();
		    options.setExportFormat(ExportFormat.XLSX);
		    options.setUseGzipCompression(false);
		    
		    URL url = reportDownloader.getDownloadUrl(options);
		    URL urlReach = reportDownloaderReach.getDownloadUrl(options);
		    URL urlGeo = reportDownloaderGeo.getDownloadUrl(options);
		    URL urlCreative = reportDownloaderCreative.getDownloadUrl(options);
		    URL urlDevices = reportDownloaderDevices.getDownloadUrl(options);
		    
		    Resources.asByteSource(url).copyTo(Files.asByteSink(file));
		    Resources.asByteSource(urlReach).copyTo(Files.asByteSink(fileReach));
		    Resources.asByteSource(urlGeo).copyTo(Files.asByteSink(fileGeo));
		    Resources.asByteSource(urlCreative).copyTo(Files.asByteSink(fileCreative));
		    Resources.asByteSource(urlDevices).copyTo(Files.asByteSink(fileDevices));
		    
		    
		   gui.pintarTextArea("Reports Downloaded successfully");
		    
		    
		    
		    gui.pintarTextArea("Creating full report... Sheet 1/6");
		    
		    FileInputStream fileIn = new FileInputStream(file);
		    FileInputStream fileInReach = new FileInputStream(fileReach);
		    FileInputStream fileInGeo= new FileInputStream(fileGeo);
		    FileInputStream fileInCreative= new FileInputStream(fileCreative);
		    FileInputStream fileInDevices= new FileInputStream(fileDevices);
		    
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFWorkbook workbookReach = new XSSFWorkbook(fileInReach);
            XSSFWorkbook workbookGeo = new XSSFWorkbook(fileInGeo);
            XSSFWorkbook workbookCreative = new XSSFWorkbook(fileInCreative);
            XSSFWorkbook workbookDevices = new XSSFWorkbook(fileInDevices);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheetReach = workbookReach.getSheetAt(0);
            XSSFSheet sheetGeo = workbookGeo.getSheetAt(0);
            XSSFSheet sheetCreative = workbookCreative.getSheetAt(0);
            XSSFSheet sheetDevices = workbookDevices.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            Iterator<Row> rowIterator2 = sheet.iterator();
           
            
            Row row = rowIterator.next();
            Row rowT = rowIterator2.next();
            rowT = rowIterator2.next();
            
            while ( rowIterator2.hasNext()) 
            {
            	row = rowIterator.next();
            	 rowT = rowIterator2.next();
            	String dayly = "";
             	 double advertiserId = 0;;
            	 String advertiser = "";
            	 double orderId = 0;
            	 String orderName = "";
            	 double adUnitId = 0;
            	 String adUnitName = "";
            	 double lineItemId = 0;
            	 String lineItemName = "";
            	 double deviceId = 0;
            	 String deviceName = "";
            	 String trafficker = "";
            	 String startDate = "";
            	 String endDate = "";
            	 String salesperson= "";
            	 double bookedCPC = 0;
            	 double bookedCPM= 0;
            	 double adServerImpressions= 0;
            	 double totalClicks= 0;
            	 Date endOrderDate1 = new Date();
            	 if(orden.getEndDateTime()!=null)
            	 {
            		 endOrderDate1 = orden.getEndDateTime().getDate();
            	 }
            	 else
            	 {
            		 endOrderDate1 = null;
            	 }
            	 
            	 
            	 double totalOrderClicks = orden.getTotalClicksDelivered();
            	 double totalOrderImpressions = orden.getTotalImpressionsDelivered();
            	 
            	 
            	 
            	 
                
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                         	
          
                	 
                	 Cell cell = cellIterator.next();
                	 
                	 
                	 for(int k = 0; cellIterator.hasNext() ;k++)
                	 {
                		 
                	if(k==0 && cell.getCellType() == cell.CELL_TYPE_STRING)
                	{
                		dayly = cell.getStringCellValue();
                	}
                		 
                		 
                		 if(k==1 && cell.getCellType() == cell.CELL_TYPE_NUMERIC )
                		 {
                			
                			 advertiserId = cell.getNumericCellValue();
                		
                		 }
                		 else if(k==2 && cell.getCellType() == cell.CELL_TYPE_STRING )
                		 {
                			
                			 advertiser = cell.getStringCellValue();
                			 
                		 }
                		 else if(k==3 && cell.getCellType() == cell.CELL_TYPE_NUMERIC)
                		 {
                			 
                			 orderId = cell.getNumericCellValue();
                			 
                		 }
                		 else if(k==4 && cell.getCellType() == cell.CELL_TYPE_STRING)
                		 {
                		
                			 
                			 orderName = cell.getStringCellValue();
                			 
                		 }
                		 else if(k==5 && cell.getCellType() == cell.CELL_TYPE_NUMERIC )
                		 {
                			
                			 adUnitId = cell.getNumericCellValue();
                		 }
                			 
                			
                		 else	 if(k==6 && cell.getCellType() == cell.CELL_TYPE_STRING )
                		 {
                			 
                			 adUnitName = cell.getStringCellValue();
                			
                		 }
                		 else if(k==7 && cell.getCellType() == cell.CELL_TYPE_NUMERIC)
                		 {
                			 
                			 lineItemId = cell.getNumericCellValue();
                			 
                		 }
                		 else if(k==8 && cell.getCellType() == cell.CELL_TYPE_STRING )
                		 {
                			
                			 lineItemName = cell.getStringCellValue();
                			
                		 }
                
                		 
                		 else if(k==9 && cell.getCellType() == cell.CELL_TYPE_STRING)
                		 {
                			 
                			 trafficker = cell.getStringCellValue();
                			 
                		 }
                		 else if(k==10 && cell.getCellType() == cell.CELL_TYPE_STRING)
                		 {
                			 
                			 startDate = cell.getStringCellValue();
                			
                		 }
                		 else if(k==11 && cell.getCellType() == cell.CELL_TYPE_STRING)
                		 {
                			 
                			 endDate= cell.getStringCellValue();
                			 
                		 }
                		 else if(k==12 && cell.getCellType() == cell.CELL_TYPE_STRING)
                		 {
                			 
                			 salesperson = cell.getStringCellValue();
                			
                		 }
                		 
                		 else if(k==13 && cell.getCellType() == cell.CELL_TYPE_NUMERIC )
                		 {
                		
                			 bookedCPC = cell.getNumericCellValue();
                			 
                		 }
                		 else if(k==14 && cell.getCellType() == cell.CELL_TYPE_NUMERIC)
                		 {
                			 
                			 bookedCPM= cell.getNumericCellValue();
                			
                		 }
                		 else if(k==15 && cell.getCellType() == cell.CELL_TYPE_NUMERIC)
                		 {
                			
                			adServerImpressions = cell.getNumericCellValue();
                		
                			k++;
                		 }
                		 
                		  if(k==16 && cell.getCellType() == cell.CELL_TYPE_NUMERIC)
                		 {
                			 
                			  cell = cellIterator.next();
                			totalClicks = cell.getNumericCellValue();
                			
                		 }
                		 if(cellIterator.hasNext())
                		 {
                		 cell = cellIterator.next();
                		 }
                		  
                	 }
                                        
                    Product producto = new Product(advertiserId, advertiser, orderId, orderName, adUnitId, adUnitName, lineItemId, lineItemName, deviceId, deviceName, trafficker, startDate, endDate, salesperson, bookedCPC, bookedCPM, adServerImpressions, totalClicks, endOrderDate1, totalOrderClicks, totalOrderImpressions , dayly);
                   
                   productsByActualOrder.add(producto);
                    
                }
		    
            FileInputStream fileOut = new FileInputStream(plantPath+"Output.xlsx");
            
          //Blank workbook
            XSSFWorkbook workbookOut = new XSSFWorkbook(fileOut); 
             
            //Create a blank sheet
            XSSFSheet sheetOut = workbookOut.getSheetAt(0);
              
//            //This data needs to be written (Object[])
//            Map<String, Object[]> data = new TreeMap<String, Object[]>();
//            data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//            data.put("2", new Object[] {1, "Amit", "Shukla"});
//            data.put("3", new Object[] {2, "Lokesh", "Gupta"});
//            data.put("4", new Object[] {3, "John", "Adwards"});
//            data.put("5", new Object[] {4, "Brian", "Schultz"});
//              
//            //Iterate over data and write to sheet
//            Set<String> keyset = data.keySet();
//            int rownum = 0;
            Product product = (Product) productsByActualOrder.get(0);
            for(int lm = 0; lm<6; lm++)
            {           	
            
            XSSFSheet sheetOutF = workbookOut.getSheetAt(lm);
            
            Row r6 = sheetOutF.getRow(5);
            Row r7 = sheetOutF.getRow(6);
            Row r8 = sheetOutF.getRow(7);
            Row r9 = sheetOutF.getRow(8);
            Row r10 = sheetOutF.getRow(9);
            
            
            Cell c6 = r6.createCell(1);
            Cell e6 = r6.createCell(4);
            
            Cell b7 = r7.createCell(1);
            Cell e7 = r7.createCell(4);

            Cell b8 = r8.createCell(1);
            Cell e8 = r8.createCell(4);
            

            Cell b9 = r9.createCell(1);
            Cell e9 = r9.createCell(4);
            

            Cell b10 = r10.createCell(1);
            
            
            
            c6.setCellValue(product.getAdvertiser());
            e6.setCellValue(product.getBookedCPC());
            b7.setCellValue(product.getAdvertiserId());
            e7.setCellValue(product.getBookedCPM());
            b8.setCellValue(product.getOrderName());
            e8.setCellValue(product.getSalesperson());
            b9.setCellValue(product.getStartDate());
            e9.setCellValue(product.getTrafficker());
            if(product.getEndOrderDate()==null)
            {
            	b10.setCellValue("-");
            }
            else
            {
            	 b10.setCellValue(product.getEndOrderDate().getYear()+"-"+ product.getEndOrderDate().getMonth()+"-"+product.getEndOrderDate().getDay());   
                 	
            }
            
            }
            
            Row r13 = sheetOut.getRow(12);
            Row r14 = sheetOut.getRow(13);
            Row r15 = sheetOut.getRow(14);
            
            Cell b13 = r13.createCell(1);
            Cell b14 = r14.createCell(1);
            Cell b15 = r15.createCell(1);
            
            b13.setCellValue(product.getTotalOrderImpressions());
            b14.setCellValue(product.getTotalOrderClicks());
            if(product.getTotalOrderClicks()==0)
            {
            	b15.setCellValue(0);
            }else
            {

                b15.setCellValue(product.getTotalOrderClicks()/product.getTotalOrderImpressions());	
            }
            
            
            
            
            int n = 20;
            int totalImpre = 0;
            int totalclicks = 0;
            for(int m = 0; m <productsByActualOrder.size(); m++)
            {
            	Product acProd = (Product) productsByActualOrder.get(m);
            	
            	totalImpre += acProd.getAdServerImpressions();
            	totalclicks += acProd.getTotalClicks();
            	
            	Row b = sheetOut.createRow(n);
            	Cell name = b.createCell(0);
            	Cell fecha = b.createCell(1);
            	Cell impresiones = b.createCell(2);
            	Cell clics = b.createCell(3);
            	Cell CTR = b.createCell(4);
            	
            	
            	name.setCellValue(acProd.getOrderName());
            	fecha.setCellValue(acProd.getDayly());
            	impresiones.setCellValue(acProd.getAdServerImpressions());
            	clics.setCellValue(acProd.getTotalClicks());
            	CTR.setCellValue(acProd.getTotalClicks()/acProd.getAdServerImpressions());
            	n++;
            }
            
            b13.setCellValue(totalImpre);
            b14.setCellValue(totalclicks);
            if(totalclicks==0)
            {
            	b15.setCellValue(0);
            }
            else
            {
            	b15.setCellValue(totalImpre/totalclicks);
            }
            
            gui.pintarTextArea("Creating full report... Sheet 2/6");
            
            //DEVICES
            //DEVICES
            XSSFSheet sheetOutDevices = workbookOut.getSheetAt(2);
            Iterator<Row> it = sheetDevices.iterator();
            
            
           
            
            int mobileRows = 22;
            int desktopRows = 22;
            
            String dateDevice = "";
            double idDevice = 0;
            double DeviceImpresionsDesktop =0;
            double DeviceImpresionsMobile =0;
            double DeviceClicksDesktop = 0;
            double DeviceClicksMobile = 0;
            double deviceImpresions = 0;
            double deviceClicks = 0;
            
            
            Row r1 = null;            
            Row firstDesktop = null;
            Cell cl = null;
            r1 = it.next();
            while(it.hasNext())
            {
            	int jeje = 0;
            	r1 = it.next();
            	Iterator<Cell> cel = r1.cellIterator();
            	while(cel.hasNext())
            	{
            		
            		cl = cel.next();
            		
            		if(jeje == 0 && cl.getCellType()==Cell.CELL_TYPE_STRING)
            		{
            			dateDevice = cl.getStringCellValue();
            		}
            		else if( jeje == 1 && cl.getCellType()==Cell.CELL_TYPE_NUMERIC)
            		{
            			idDevice = cl.getNumericCellValue();
            			
            		}
            		else if( jeje == 3 && cl.getCellType() == Cell.CELL_TYPE_NUMERIC)
            		{
            			deviceImpresions = cl.getNumericCellValue();
            		}
            		else if( jeje == 4 && cl.getCellType() == Cell.CELL_TYPE_NUMERIC)
            		{
            			deviceClicks = cl.getNumericCellValue();
            		}
            		
            		   		
            	
            	
            }
            	if( idDevice == 30000.0)
        		{
        			DeviceClicksDesktop+= deviceClicks;
        			DeviceImpresionsDesktop += deviceImpresions;
        			if(sheetOutDevices.getRow(desktopRows)!= null)
        			{
        				firstDesktop = sheetOutDevices.getRow(desktopRows);
        			}
        			else
        			{
        				 firstDesktop = sheetOutDevices.createRow(desktopRows);
        			}
        			
        			Cell day = firstDesktop.createCell(0);
        			Cell imp= firstDesktop.createCell(1);
        			Cell clics = firstDesktop.createCell(2);
        			Cell ctr = firstDesktop.createCell(3);
        			if(dateDevice !="Total")
        			{
        				day.setCellValue(dateDevice);
            			imp.setCellValue(deviceImpresions);
            			clics.setCellValue(deviceClicks);
        			}
        			
        			if(deviceClicks!=0)
        			{
        				ctr.setCellValue(deviceClicks/deviceImpresions);
        			}
        			else
        			{
        				ctr.setCellValue(0);
        			}
        			
        			desktopRows++;
        		}
        		
        		else
        		{
        		DeviceClicksMobile += deviceClicks;
        		DeviceImpresionsMobile += deviceImpresions;
        		if(sheetOutDevices.getRow(mobileRows)!= null)
    			{
    				firstDesktop = sheetOutDevices.getRow(mobileRows);
    			}
    			else
    			{
    				 firstDesktop = sheetOutDevices.createRow(mobileRows);
    			}
        			
        		Cell day = firstDesktop.createCell(4);
    			Cell imp= firstDesktop.createCell(5);
    			Cell clics = firstDesktop.createCell(6);
    			Cell ctr = firstDesktop.createCell(7);
    			if(dateDevice !="Total")
    			{
    				day.setCellValue(dateDevice);
        			imp.setCellValue(deviceImpresions);
        			clics.setCellValue(deviceClicks);
    			}
    			
    			if(deviceClicks!=0)
    			{
    				ctr.setCellValue(deviceClicks/deviceImpresions);
    			}
    			else
    			{
    				ctr.setCellValue(0);
    			}
    			
        		mobileRows++;
        		
        		}
        		jeje++;
        		System.out.println(jeje);
        	}
        	
        
            Row deviceimp = sheetOutDevices.getRow(14);
            Row deviceclicks = sheetOutDevices.getRow(15);
            Row devicectr = sheetOutDevices.getRow(16);
            
            Cell devimpresions = deviceimp.createCell(1);
            Cell devimpresions2 = deviceimp.createCell(3);
            Cell devimpresions3 = deviceclicks.createCell(1);
            Cell devimpresions4 = deviceclicks.createCell(3);
            Cell devimpresions5 = devicectr.createCell(1);
            Cell devimpresions6 = devicectr.createCell(3);
            
            devimpresions.setCellValue(DeviceImpresionsDesktop);
            devimpresions2.setCellValue(DeviceImpresionsMobile);
            devimpresions3.setCellValue(DeviceClicksDesktop);
            devimpresions4.setCellValue(DeviceClicksMobile);
            if(DeviceClicksDesktop==0)
            {
            	devimpresions5.setCellValue(0);
            }
            else
            {
            	  devimpresions5.setCellValue(DeviceClicksDesktop/DeviceImpresionsDesktop);
            }
            
            if(DeviceClicksMobile==0)
            {
            	devimpresions6.setCellValue(0);
            }
            else
            devimpresions6.setCellValue(DeviceClicksMobile/DeviceImpresionsMobile);
            
            
            
            
            
           
            gui.pintarTextArea("Creating full report... Sheet 3/6");
            // REACH
            int im = 1;
            double reach = 0;
            double reachf = 0;
            while(sheetReach.getRow(im)!= null && sheetReach.getRow(im).getCell(0).getCellType()!=Cell.CELL_TYPE_STRING)
            {
            	Row m = sheetReach.getRow(im);
            	Cell un = m.getCell(4);
            	Cell av = m.getCell(5);
            	if(un.getCellType()==Cell.CELL_TYPE_NUMERIC)
            	{
            		reach = un.getNumericCellValue();
            	}
            	
            	if(av.getCellType()== Cell.CELL_TYPE_NUMERIC)
            	{
            		reachf = av.getNumericCellValue();
            	}
            	
            	
            	im++;
            }
            XSSFSheet sheetOutReach = workbookOut.getSheetAt(1);
            
            Row reach1 =sheetOutReach.getRow(13);
            Row reach2 =sheetOutReach.getRow(14);
            Cell reach11 = reach1.createCell(1);
            Cell reach22 = reach2.createCell(1);
            reach11.setCellValue(reach);
            reach22.setCellValue(reachf);
           
            //GEOGRAPH
            
            gui.pintarTextArea("Creating full report... Sheet 4/6");
            XSSFSheet sheetOutGeo = workbookOut.getSheetAt(4);
            int am = 1;
            String ciudad ="";
            double impressions = 0;
            double geoclics = 0;
            double totalImpGeo = 0;
            double totalclicGeo = 0;
            int outrow = 21;
            while(sheetGeo.getRow(am)!= null && !sheetGeo.getRow(am).getCell(0).getStringCellValue().equals("Total"))
            {
            	
            	Row m = sheetGeo.getRow(am);
            	Cell cit = m.getCell(0);
            	Cell mm = m.getCell(6);
            	Cell nn = m.getCell(7);
            	ciudad = cit.getStringCellValue();
            	impressions = mm.getNumericCellValue();
            	totalImpGeo += impressions;
            	geoclics = nn.getNumericCellValue();
            	totalclicGeo += geoclics;
            	am++;
            	
            	Row write = sheetOutGeo.createRow(outrow);
            	Cell cita = write.createCell(0);
            	Cell citaImp = write.createCell(1);
            	Cell citaclic = write.createCell(2);
            	Cell ctrclicgeo = write.createCell(3);
            	
            	cita.setCellValue(ciudad);
            	citaImp.setCellValue(impressions);
            	citaclic.setCellValue(geoclics);
            	if(geoclics==0)
            	{
            		ctrclicgeo.setCellValue(0);
            	}
            	else
            	{
            		ctrclicgeo.setCellValue(impressions/geoclics);
            	}
            	outrow++;
            }
            
            Row tImp = sheetOutGeo.getRow(13);
            Cell tim = tImp.createCell(1);
            tim.setCellValue(totalImpGeo);
            Row tImp2 = sheetOutGeo.getRow(14);
            Cell top = tImp2.createCell(1);
            top.setCellValue(totalclicGeo);
            Row tImp3= sheetOutGeo.getRow(15);
            Cell tep = tImp3.createCell(1);
            if(totalclicGeo==0)
            {
            	 tep.setCellValue(0);
                 
            }
            else
            {
            	tep.setCellValue(totalImpGeo/totalclicGeo);
            }
            
            gui.pintarTextArea("Creating full report... Sheet 5/6");
            //ANUNCIOS - CREATIVE
            XSSFSheet sheetOutCreative = workbookOut.getSheetAt(5);
            
            int ba = 1;
            String creativ = "";
            double creatimpre = 0;
           double creatclics = 0;
            double totalcreatimpre = 0;
            double totalcreatclic = 0;
            double ctrcreat = 0;
            int creativerow = 21;
            while(sheetCreative.getRow(ba)!= null && sheetCreative.getRow(ba).getCell(0).getCellType()!=Cell.CELL_TYPE_STRING)
            {
            	Row m = sheetCreative.getRow(ba);
            	Cell ca= m.getCell(1);
            	Cell ac = m.getCell(3);
            	Cell cac = m.getCell(4);
            	if(ca.getCellType()==Cell.CELL_TYPE_STRING)
            	{
            		creativ = ca.getStringCellValue();
            	}
            	else
            	{
            		creativ =""+ ca.getNumericCellValue();
            	}
            	creatimpre = ac.getNumericCellValue();
            	totalcreatimpre+=creatimpre;
            	creatclics = cac.getNumericCellValue();
            	totalcreatclic += creatclics;
            	
            	Row creativeFirstRow = sheetOutCreative.createRow(creativerow);
            	Cell firstA = creativeFirstRow.createCell(0);
            	Cell firstB = creativeFirstRow.createCell(1);
            	Cell firstC = creativeFirstRow.createCell(2);
            	Cell firstD = creativeFirstRow.createCell(3);
            	
            	firstA.setCellValue(creativ);
            	firstB.setCellValue(creatimpre);
            	firstC.setCellValue(creatclics);
            	if(creatclics==0)
            	{
            		firstD.setCellValue(0);
            	}
            	else
            	{
            		firstD.setCellValue(creatimpre/creatclics);
            	}
            	
            	
            	creativerow++;
            	ba++;
            	
            }
            if(totalcreatclic == 0)
        	{
        		ctrcreat = 0;
        	}
        	else
        	{
        		ctrcreat = totalcreatimpre/totalcreatclic;
        	}
            
            Row extout = sheetOutCreative.getRow(13);
            Cell u1 = extout.createCell(1);
            u1.setCellValue(totalcreatimpre);
            Row extout2 = sheetOutCreative.getRow(14);
            Cell u2 = extout2.createCell(1);
            u2.setCellValue(totalcreatclic);
            Row extout3 = sheetOutCreative.getRow(15);
            Cell u3 = extout3.createCell(1);
            if(totalcreatclic == 0)
            {
            	u3.setCellValue(0);
            }
            else
            {
            	u3.setCellValue(totalcreatimpre/totalcreatclic);
            }
            gui.pintarTextArea("Creating full report... Sheet 6/6");
            
            
            //Ubicacion
            
            XSSFSheet sheetOutUnit = workbookOut.getSheetAt(3);
            
            double adunittotalimpre = 0;
            double adunittotalclicks = 0;
            int actrow = 21;
            for(int pi = 0; pi< productsByActualOrder.size(); pi++)
            {
            	Product actual = (Product) productsByActualOrder.get(pi);
            	String adunitname = actual.getAdUnitName();
            	double adunitid = actual.getAdUnitId();
            	double adunitimpre = actual.getAdServerImpressions();
            	double adunitclicks = actual.getTotalClicks();
            	adunittotalclicks += adunitclicks;
            	adunittotalimpre += adunitimpre;
            	double ctr =0;
            	if(adunitimpre!=0)
            	{
            		 ctr = adunitclicks/adunitimpre;
            	}
            	else
            	{
            		ctr = 0;
            	}
            	
            	Row act = sheetOutUnit.createRow(actrow);
            	Cell name = act.createCell(0);
            	Cell imp = act.createCell(1);
            	Cell clicks = act.createCell(2);
            	Cell ctrI  = act.createCell(3);
            	
            	name.setCellValue(adunitname);
            	imp.setCellValue(adunitimpre);
            	clicks.setCellValue(adunitclicks);
            	ctrI.setCellValue(ctr);
            	actrow++;
            }
            
            Row p1 = sheetOutUnit.getRow(13);
            Row p2 = sheetOutUnit.getRow(14);
            Row p3 = sheetOutUnit.getRow(15);
            
            Cell p1p1 = p1.createCell(1);
            Cell p1p2 = p2.createCell(1);
            Cell p1p3 = p3.createCell(1);
            
            
            p1p1.setCellValue(adunittotalimpre);
            p1p2.setCellValue(adunittotalclicks);
            
            if(adunittotalimpre==0)
            {
            	p1p3.setCellValue(0);
            }
            else
            {
            	p1p3.setCellValue(adunittotalclicks/adunittotalimpre);
            }
            
           
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
          //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(parentPath+"ReporteFinal-"+orden.getName()+".xlsx");
            workbookOut.write(out);
            out.close();
            gui.pintarTextArea(orden.getName()+".xlsx written successfully on disk.");
         
          
	    }
	    
            
	    
	    
            }
			

	    	    	
	public ArrayList deduplicateArray()
	{
		
		for(int i = 0; i< orderId.size();i++)
		{
			for(int j = i+1; j< orderId.size();j++)
			{
				Order search = (Order) orderId.get(i);
				Order comp = (Order) orderId.get(j);
				if(search.getId()== comp.getId())
				{
					orderId.remove(comp);
				}
			}
		}
		return orderId;
	}

	private void updateTextArea(final String text ) {
		  SwingUtilities.invokeLater(new Runnable() {
		    public void run() {
		      textArea.append(text);
		      textArea.repaint();
		      textArea.revalidate();
		    }
		  });
		}
		 
		private void redirectSystemStreams() {
		  OutputStream out = new OutputStream() {
		    @Override
		    public void write(int b) throws IOException {
		      updateTextArea(String.valueOf((char) b));
		    }
		 
		    @Override
		    public void write(byte[] b, int off, int len) throws IOException {
		      updateTextArea(new String(b, off, len));
		    }
		 
		    @Override
		    public void write(byte[] b) throws IOException {
		      write(b, 0, b.length);
		    }
		  };
		 
		  System.setOut(new PrintStream(out, true));
		  System.setErr(new PrintStream(out, true));
		}

		private String getYesterdayDateString() {
	        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
	        Calendar cal = Calendar.getInstance();
	        cal.add(Calendar.DATE, -1);    
	        return dateFormat.format(cal.getTime());
	}
	public  void readProp(TextArea textArea , File p)
	{
		InputStream inputStream;
		try
		{
			
			
			Properties prop = new Properties();
			
			inputStream = new FileInputStream(p);
			if(inputStream !=null)
			{
				prop.load(inputStream);
			}
			else
			{
				throw new FileNotFoundException("Property file: "+PROPERTIES+" Not found.");
			}
			
			//GET PROPERTIES
			
			
			parentPath = prop.getProperty("ParentPath");
			
			reportPath = prop.getProperty("ReportPath");
			
			plantPath = prop.getProperty("PlantPath");
			
			gui.pintarTextArea("Properties successfully read.");
			gui.pintarTextArea(parentPath);
			gui.pintarTextArea(reportPath);
			gui.pintarTextArea(plantPath);
			
			
		textArea.repaint();
		textArea.revalidate();
			
			
		}
		catch(Exception e)
		{
			e.getMessage();
			e.printStackTrace();
		}
	}
	
	
	  public static void main(String[] args) throws Exception {
	    // Generate a refreshable OAuth2 credential.
		
			  
		
	    Credential oAuth2Credential = new OfflineCredentials.Builder()
	  
	        .forApi(Api.DFP)
	        .fromFile()
	        .build()
	        .generateCredential();

	
	    
	    
	    // Construct a DfpSession.
	    DfpSession session = new DfpSession.Builder()
	        .fromFile()
	        .withOAuth2Credential(oAuth2Credential)
	        .build();

	    DfpServices dfpServices = new DfpServices();

//	    runExample(dfpServices, session);
//	    runExampleReport(dfpServices, session);
//	    DFPCaracolWorld caracol = new DFPCaracolWorld(dfpServices, session);
	  }
	}






