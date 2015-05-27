import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import jxl.CellView;
import jxl.Workbook;
import jxl.write.*;


public class VerizonAvailability {
	
	public static void main(String[] args) throws IOException, WriteException{
		
		Calendar now = new GregorianCalendar();
		SimpleDateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy");
		
		String date = dateFormat.format(now.getTime());
		
		Crawler availCrawler = new Crawler("C:\\Users\\User\\Documents\\Dymax Crawler\\Availability Sheet.txt");
		
		ArrayList<String> availList = availCrawler.getSKUList();
		
		availCrawler.signIn("https://indirectorders.verizonwireless.com/xt_webloginvalidation.aspx", 
				args[0], args[1]);
		
		int row = 2;
		final int descCol = 0;
		final int priceCol = 1;
		final int qtyCol = 2;
		
		WritableWorkbook myBook = Workbook.createWorkbook(new File("C:\\Users\\User\\Documents\\Availability\\Verizon"
				+ " Availability " + date + ".xls"));
		
		WritableSheet sheet = myBook.createSheet("Availability", 0);
		Label descriptionCell, qtyCell, priceCell;
		
		for(String sku : availList){
			int qtyAvail = availCrawler.checkQty(sku);
			String description = availCrawler.getDescription();
			String price = availCrawler.getPrice();
			
			descriptionCell = new Label(descCol, row, description);
			qtyCell = new Label(qtyCol, row, ""+qtyAvail);
			priceCell = new Label(priceCol, row, price);
			
			sheet.addCell(descriptionCell);
			sheet.addCell(qtyCell);
			sheet.addCell(priceCell);
			
			row++;
		}
		
	    CellView cell = sheet.getColumnView(descCol);
	    cell.setAutosize(true);
	    sheet.setColumnView(descCol, cell);

		
		myBook.write();
		myBook.close();
		availCrawler.quit();
		
	}
}
