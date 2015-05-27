import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;

import jxl.CellType;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;


public class IconicCrawler extends Crawler{

	private String customerName, shipToAddress1, shipToAddress2, shipToCity, trackingNumber, preOrderNumber;
	private ArrayList<IconicList> storeCodes = IconicList.generateList();
	private final int PREORDER_NUMBER_COLUMN = 8, TRACKING_COLUMN = 9, CUSTOMER_NAME_COLUMN = 10;
	private String testShippingCode;
	private Workbook inputBook;
	private WritableWorkbook outputBook;
	private WritableSheet sheet;
		
	public void setupExcel(String input, String output) throws BiffException, IOException{
		File inputFile = new File(input);
		File outputFile = new File(output);

		inputBook = Workbook.getWorkbook(inputFile);
		outputBook = Workbook.createWorkbook(outputFile, inputBook);
		sheet = outputBook.getSheet(0);
	}
	
	public String readPreorderNumber(int row){
		WritableCell sourceCell = sheet.getWritableCell(PREORDER_NUMBER_COLUMN, row);
		
		preOrderNumber = sourceCell.getContents();
		return preOrderNumber;
	}
	
	public void writePreorderInfo(int row) throws RowsExceededException, WriteException{
		Label trackingNumber = new Label(TRACKING_COLUMN, row, getTrackingNumber());
		sheet.addCell(trackingNumber);
		
		Label customerName = new Label(CUSTOMER_NAME_COLUMN, row, getCustomerName());
		sheet.addCell(customerName);
	}
	
	public void compileSheet() throws RowsExceededException, WriteException{
		int rowCounter = 0;
		String oldPreOrder = "";
		while(!readPreorderNumber(rowCounter).equals("")){
			
			String preOrder =readPreorderNumber(rowCounter);
			
			if(preOrder.equals(oldPreOrder)){
				writePreorderInfo(rowCounter);
			}else{
				compilePreOrder(preOrder);
				writePreorderInfo(rowCounter);
			}	
			
			rowCounter++;
			
			oldPreOrder = preOrder;
		}	
	}
	
	public void compilePreOrder(String preorderNumber){
		try{
			navigateToPreorder(preorderNumber);
			findCustomerName();
			findTrackingNumber();
			findShippingAddress();
			
			if(!shippedToStore(getTestShippingCode(), getShipToAddress1())){
				this.customerName = getCustomerName() + " " +getShipToAddress1() + " " +
									getShipToAddress2() + " " + getShipToCity();
			}
		}catch(Exception e){
			try{
				String newPreorderNumber = preorderNumber.substring(0, preorderNumber.length()-1) + 1;
			
				navigateToPreorder(newPreorderNumber);
				findCustomerName();
				findTrackingNumber();
				findShippingAddress();
			
				if(!shippedToStore(getTestShippingCode(), getShipToAddress1())){
					this.customerName = getCustomerName() + "\n" + getShipToAddress1() + "\n"+
									getShipToAddress2() +"\n" + getShipToCity();
				}
			}catch(Exception f){
				customerName = "Seach failed, look up manually.";
				trackingNumber = "Seach failed, look up manually.";
				shipToAddress1 = "Seach failed, look up manually.";
			}	
		}
	}
	
	/**
	 * Opens a specific Iconic Order based on its preorder number
	 * 
	 * @param preorderNumber the preorder number as provided by Verizon
	 */
	public void navigateToPreorder(String preorderNumber){
		getDriver().get("https://indirectorders.verizonwireless.com/xt_documentsearch.aspx");
		
		WebElement dropDown = getDriver().findElement(By.id("ctl00_MainPlaceHolder_ddlCriteriaMain"));
		Select preorder = new Select(dropDown);
		preorder.selectByVisibleText("Pre-Order Number");
		
		//trying to find the preorder text field without a delay was causing issues
		//since it takes a moment for the field to populate after clicking preorder
		try {
			Thread.sleep(500);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		
		getDriver().findElement(By.name("ctl00$MainPlaceHolder$txtFromCriteriaNoDate")).sendKeys(preorderNumber);
		getDriver().findElement(By.id("ctl00_MainPlaceHolder_btnSearch")).click();
		
		//we need to grab the store's shipping code before going into the actual order
		//so that we can verify that the phone is being shipped to the store and not
		//a customer's house
		this.testShippingCode = getDriver().findElement(By.xpath("//*[@id=\"ctl00_MainPlaceHolder_grdOrdersList\"]/tbody/tr[2]/td[5]")).getText();
		
		getDriver().findElement(By.id("ctl00_MainPlaceHolder_grdOrdersList_ctl02_LinkButton1")).click();
	}
	
	public void findCustomerName(){
		WebElement custName = getDriver().findElement(By.id("ctl00_MainPlaceHolder_lblShipToContact"));
		
		this.customerName = custName.getText();

	}
	
	public void findTrackingNumber(){
		WebElement shipTo = getDriver().findElement(By.id("ctl00_MainPlaceHolder_lblTracking"));
		
		this.trackingNumber = shipTo.getText();
	}
	
	public void findShippingAddress(){
		WebElement custAdd = getDriver().findElement(By.id("ctl00_MainPlaceHolder_lblShipToAddress1"));
		WebElement custAdd2 = getDriver().findElement(By.id("ctl00_MainPlaceHolder_lblShipToAddress2"));
		WebElement city = getDriver().findElement(By.id("ctl00_MainPlaceHolder_lblShipToCityStZip"));
		
		this.shipToAddress1 = custAdd.getText();
		this.shipToAddress2 = custAdd2.getText();
		this.shipToCity = city.getText();
	}
	
	public ArrayList<IconicList> getStoreCodes(){
		return storeCodes;
	}
	
	public Boolean shippedToStore(String testShippingCode, String shipToAddress1){
		this.shipToAddress1 = shipToAddress1;
		for(IconicList e : storeCodes){
			if(e.getCode().equals(testShippingCode)){
				
				if(e.getCode().equals("A53")){//weird corner case store that is loaded in Dymax oddly
					return shipToAddress1.startsWith(e.getAddress()) || shipToAddress1.startsWith("DBA");
				}else{
					return shipToAddress1.startsWith(e.getAddress());
				}	
			}
		}
		return false;
	}
	
	public String getCustomerName() {
		return customerName;
	}
	public String getShipToAddress1() {
		return shipToAddress1;
	}
	public String getShipToAddress2() {
		return shipToAddress2;
	}
	public String getShipToCity() {
		return shipToCity;
	}
	public String getTrackingNumber() {
		return trackingNumber;
	}
	public String getPreOrderNumber(){
		return preOrderNumber;
	}
	public String getTestShippingCode() {
		return testShippingCode;
	}
	public Workbook getInputBook() {
		return inputBook;
	}
	public WritableWorkbook getOutputBook() {
		return outputBook;
	}

	public void write(WritableWorkbook outputBook) throws IOException, WriteException{
		outputBook.write();
		outputBook.close();
	}

	public static void main(String[] args) throws BiffException, IOException, WriteException{
		final long startTime = System.currentTimeMillis();
		
		Calendar now = new GregorianCalendar();
		SimpleDateFormat dateFormat = new SimpleDateFormat("MMddyyyy");
		String date = dateFormat.format(now.getTime());
		String outputFile = "C:/Users/User/Documents/Iconic Orders/Iconic Sheet "+date+" " + ".xls";
		IconicCrawler myCrawl = new IconicCrawler();
		
		myCrawl.setupExcel(args[2], outputFile);

		myCrawl.signIn("https://indirectorders.verizonwireless.com/xt_webloginvalidation.aspx", 
				args[0], args[1]);
		
		myCrawl.compileSheet();
		
		myCrawl.quit();
		myCrawl.write(myCrawl.getOutputBook());
		
		final long endTime = System.currentTimeMillis();
		
		float timeElapsed = (endTime-startTime);
		
		System.out.println("total execution time: " + timeElapsed);
	
	}
	
}
