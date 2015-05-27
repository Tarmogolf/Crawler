import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import twitter4j.Twitter;
import twitter4j.TwitterException;
import twitter4j.TwitterFactory;

/**
 * Crawls through the Verizon Dymax website, checking
 * the availability of a user specified list of SKUs.
 * Notifies the user via Twitter if something becomes
 * available.
 * 
 * @author Stan Bessey
 */
public class Crawler {

	private WebDriver driver; //the type of browser to be used. Firefox has the fewest/easiest dependencies
	private WebElement userName, password; //the username and password fields on the Dymax website
	private final String baseURL = "https://indirectorders.verizonwireless.com/xt_product.aspx?sku=";
	private String description, price;
	private Twitter twitter;
	private final String users[] = {"@MBFInventory","@skress31"}; //list of people to tweet to when items show up
	private Calendar now = new GregorianCalendar();
	private SimpleDateFormat dateFormat = new SimpleDateFormat("hh:mm a");
	private ArrayList<String> SKUlist;


	/**
	 * Default constructor methods. Could add additional in the future 
	 * if the user prefers to use Chrome or IE.
	 * @throws FileNotFoundException 
	 */
	public Crawler(String fileList) throws FileNotFoundException{
		this.driver = new FirefoxDriver();
		this.twitter = TwitterFactory.getSingleton();

		File file = new File(fileList);
		this.SKUlist = generateSKUList(file);

	}
	
	public Crawler(){
		this.driver = new FirefoxDriver();
		this.twitter = TwitterFactory.getSingleton();
	}

	/**
	 * Launches the web browser to the given URL and logs in
	 * with the user specified ID and password for the site.
	 *
	 * @param URL - the website to navigate to
	 * @param ID - the user's ID for the Dymax site
	 * @param pw - the user's password for the Dymax site.
	 */
	public void signIn(String URL, String ID, String pw){
		//launch the webpage
		driver.get(URL);

		//find the username and password fields
		this.userName = driver.findElement(By.name("ctl00$MainPlaceHolder$TextBox1"));
		this.password = driver.findElement(By.name("ctl00$MainPlaceHolder$TextBox2"));

		//send the username and password to the system
		userName.sendKeys(ID);
		password.sendKeys(pw);

		//click the submit button
		WebElement submitButton = driver.findElement(By.name("ctl00$MainPlaceHolder$Button1"));
		submitButton.click();
	}

	/**
	 * Checks the Dymax catalog to see how many of the
	 * specified SKU are available.
	 * 
	 * @param partNumber the SKU that should be 
	 * @return the number of items available for the specified SKU
	 */
	public int checkQty(String partNumber){
		String urlToCheck = baseURL + partNumber;

		driver.get(urlToCheck);

		WebElement qtyField = driver.findElement(By.id("ctl00_MainPlaceHolder_lblAvailQty"));
		WebElement descriptField = driver.findElement(By.id("ctl00_MainPlaceHolder_lblItemDescription"));
		WebElement priceField = driver.findElement(By.id("ctl00_MainPlaceHolder_lblUnitPrice"));

		this.description = descriptField.getText();
		this.price = priceField.getText();
		int qtyAvail = Integer.parseInt(qtyField.getText());

		return qtyAvail;
	}

	/**
	 * Reads a plaintext file for the SKUs that should be checked
	 * on Dymax's catalog.
	 * 
	 * @param file The file to be scanned for part numbers
	 * @return ArrayList of all of the part numbers in the given file
	 * @throws FileNotFoundException Thrown if the given file cannot be located
	 */
	public ArrayList<String> generateSKUList(File file) throws FileNotFoundException{
		Scanner scan = new Scanner(file);
		ArrayList<String> SKUList = new ArrayList<String>();

		while(scan.hasNext()){
			SKUList.add(scan.next().trim());
		}
		scan.close();

		return SKUList;
	}

	public void sendTweet(String description, int qty) throws TwitterException{
		String tweet = "";
		for(String s : users){
			tweet+=s+ " ";
		}
		tweet += " "+qty+" "+description;

		twitter.updateStatus(tweet);
	}

	public void sendTweet(String msg) throws TwitterException{
		twitter.updateStatus(msg);
	}

	public void crawl(ArrayList<String> SKUList) throws TwitterException{		
		ArrayList<String> foundList = new ArrayList<String>();
		//check the quantity on each SKU in the list
		for(int i = 0; i < SKUList.size(); i++){
			String date = dateFormat.format(now.getTime());
			int qtyAvail = checkQty(SKUList.get(i)); //must be called before getDescription
			String desc = getDescription();
			if(qtyAvail > 0){
				sendTweet(desc, qtyAvail);
				foundList.add(SKUList.get(i));
			}else{
				System.out.println(desc + " not avail @ " + date);
			}
		}
		//remove SKUs that have been found from the list to check
		//cannot do within the above loop to avoid modification
		//issues
		for(String i : foundList){
			SKUList.remove(i);
		}
		//insert a delay
		try {
			Thread.sleep(90000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		//update the current time
		now = GregorianCalendar.getInstance();
	}

	/**
	 * Closes the web browser
	 */
	public void quit(){
		driver.quit();
	}

	public WebDriver getDriver(){
		return driver;
	}

	public String getDescription(){
		return description;
	}

	public String getPrice(){
		return price;
	}

	public ArrayList<String> getSKUList(){
		return SKUlist;
	}

	public static void main(String[] args) throws FileNotFoundException, TwitterException, BiffException, IOException, WriteException{
		//sanity check to make sure username and password have been specified
		if(args.length != 2){
			System.out.println("Must pass exactly two arguments. First should be the user ID for Dymax, "
					+ "the second should be your password for Dymax.");
			System.exit(1);
		}

		Calendar now = new GregorianCalendar();

		Crawler dymaxCrawler = new Crawler("C:\\Users\\User\\Documents\\Dymax Crawler\\Dymax Crawler.txt");

		//open web browser to specified URL and log in
		dymaxCrawler.signIn("https://indirectorders.verizonwireless.com/xt_webloginvalidation.aspx", 
				args[0], args[1]);

		//run the program until 9:00 pm and the list has items to check for
		while(now.get(Calendar.HOUR_OF_DAY)< 21 && dymaxCrawler.getSKUList().size() != 0){
			dymaxCrawler.crawl(dymaxCrawler.getSKUList());
		}

		dymaxCrawler.quit();
	}
}
