import os
import time
import random
import pandas as pd
import smtplib
import schedule
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime

class MCPServer:
    def __init__(self):
        # Configuration settings
        self.data_file = "leads_database.xlsx"
        self.daily_target = 50
        self.search_queries = [
            "manufacturing companies in Bengaluru",
            "industrial factories in Bengaluru",
            "manufacturing brands in Bangalore",
            "production facilities Bengaluru",
            "manufacturing units in Bangalore"
        ]
        self.email_config = {
            "sender_email": "your_email@gmail.com",
            "sender_password": "your_app_password",  # Use app password for Gmail
            "recipient_email": "your_email@gmail.com",
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587
        }
        
        # Initialize the database if it doesn't exist
        self.initialize_database()
        
    def initialize_database(self):
        """Create or load the Excel database"""
        if not os.path.exists(self.data_file):
            df = pd.DataFrame(columns=[
                'Business Name', 'Address', 'Phone', 'Website', 
                'Category', 'Rating', 'Reviews', 'Date Added', 'Contacted'
            ])
            df.to_excel(self.data_file, index=False)
            print(f"Created new database: {self.data_file}")
        else:
            print(f"Using existing database: {self.data_file}")
    
    def setup_webdriver(self):
        """Configure and return a Chrome webdriver instance"""
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Run in background
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36")
        
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    
    def scrape_google_maps(self):
        """Scrape business data from Google Maps"""
        print("Starting Google Maps scraping process...")
        driver = self.setup_webdriver()
        collected_data = []
        
        # Load existing data to avoid duplicates
        existing_df = pd.read_excel(self.data_file)
        existing_names = set(existing_df['Business Name'].str.lower())
        
        for query in self.search_queries:
            if len(collected_data) >= self.daily_target:
                break
                
            search_url = f"https://www.google.com/maps/search/{query.replace(' ', '+')}"
            driver.get(search_url)
            
            # Wait for results to load
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='feed']"))
                )
            except TimeoutException:
                print(f"Timeout waiting for results for query: {query}")
                continue
                
            # Scroll to load more results
            for _ in range(5):  # Scroll 5 times
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", 
                                     driver.find_element(By.CSS_SELECTOR, "div[role='feed']"))
                time.sleep(1)
                
            # Extract business listings
            business_elements = driver.find_elements(By.CSS_SELECTOR, "div[role='article']")
            
            for element in business_elements:
                if len(collected_data) >= self.daily_target:
                    break
                    
                try:
                    # Click on the listing to view details
                    element.click()
                    time.sleep(2)  # Wait for details to load
                    
                    # Extract business information
                    name = driver.find_element(By.CSS_SELECTOR, "h1.fontHeadlineLarge").text
                    
                    # Skip if already in database
                    if name.lower() in existing_names:
                        continue
                        
                    # Extract more details
                    try:
                        address = driver.find_element(By.CSS_SELECTOR, "button[data-item-id='address']").text
                    except NoSuchElementException:
                        address = "Not available"
                        
                    try:
                        phone = driver.find_element(By.CSS_SELECTOR, "button[data-item-id^='phone:']").text
                    except NoSuchElementException:
                        phone = "Not available"
                        
                    try:
                        website = driver.find_element(By.CSS_SELECTOR, "a[data-item-id='authority']").get_attribute("href")
                    except NoSuchElementException:
                        website = "Not available"
                        
                    try:
                        category = driver.find_element(By.CSS_SELECTOR, "button[jsaction='pane.rating.category']").text
                    except NoSuchElementException:
                        category = "Not available"
                        
                    try:
                        rating = driver.find_element(By.CSS_SELECTOR, "div.fontDisplayLarge").text
                        reviews = driver.find_element(By.CSS_SELECTOR, "div.fontBodyMedium > span").text.replace(" reviews", "")
                    except NoSuchElementException:
                        rating = "Not available"
                        reviews = "0"
                    
                    # Add to collected data
                    collected_data.append({
                        'Business Name': name,
                        'Address': address,
                        'Phone': phone,
                        'Website': website,
                        'Category': category,
                        'Rating': rating,
                        'Reviews': reviews,
                        'Date Added': datetime.now().strftime("%Y-%m-%d"),
                        'Contacted': 'No'
                    })
                    
                    # Add to existing names to avoid duplicates
                    existing_names.add(name.lower())
                    
                    # Random delay to avoid detection
                    time.sleep(random.uniform(1, 3))
                    
                except Exception as e:
                    print(f"Error processing listing: {e}")
                    continue
        
        driver.quit()
        return collected_data
    
    def update_database(self, new_data):
        """Add new leads to the Excel database"""
        if not new_data:
            print("No new data to add to database")
            return False
            
        print(f"Adding {len(new_data)} new leads to database")
        existing_df = pd.read_excel(self.data_file)
        new_df = pd.DataFrame(new_data)
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        updated_df.to_excel(self.data_file, index=False)
        return True
    
    def generate_daily_report(self):
        """Generate a report of today's new leads"""
        df = pd.read_excel(self.data_file)
        today = datetime.now().strftime("%Y-%m-%d")
        today_leads = df[df['Date Added'] == today]
        
        if today_leads.empty:
            return None
            
        report_file = f"leads_report_{today}.xlsx"
        today_leads.to_excel(report_file, index=False)
        return report_file
    
    def send_email_notification(self, report_file):
        """Send email with the daily leads report"""
        if not report_file:
            print("No report file to send")
            return False
            
        try:
            # Set up email
            msg = MIMEMultipart()
            msg['From'] = self.email_config['sender_email']
            msg['To'] = self.email_config['recipient_email']
            msg['Subject'] = f"Daily Leads Report - {datetime.now().strftime('%Y-%m-%d')}"
            
            body = f"""
            Hello,
            
            Attached is today's lead generation report with {self.daily_target} manufacturing/brand leads from Bengaluru.
            These leads are suitable for pitching ACVISS products.
            
            This is an automated email from your MCP Server.
            """
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach the report file
            attachment = open(report_file, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {report_file}")
            msg.attach(part)
            attachment.close()
            
            # Connect to server and send email
            server = smtplib.SMTP(self.email_config['smtp_server'], self.email_config['smtp_port'])
            server.starttls()
            server.login(self.email_config['sender_email'], self.email_config['sender_password'])
            server.send_message(msg)
            server.quit()
            
            print(f"Email notification sent with report: {report_file}")
            return True
            
        except Exception as e:
            print(f"Error sending email: {e}")
            return False
    
    def run_daily_process(self):
        """Execute the complete daily lead generation process"""
        print(f"\n--- Starting MCP Daily Process: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
        
        # Step 1: Scrape data from Google Maps
        new_leads = self.scrape_google_maps()
        
        # Step 2: Update the database
        self.update_database(new_leads)
        
        # Step 3: Generate daily report
        report_file = self.generate_daily_report()
        
        # Step 4: Send email notification
        if report_file:
            self.send_email_notification(report_file)
            
        print(f"--- Finished MCP Daily Process: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
    
    def start_scheduler(self):
        """Start the scheduler to run the process daily"""
        schedule.every().day.at("08:00").do(self.run_daily_process)
        
        print("MCP Server started. Will generate leads daily at 08:00")
        print("Press Ctrl+C to stop the server")
        
        # Run immediately for the first time
        self.run_daily_process()
        
        # Keep the scheduler running
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute

# Main execution
if __name__ == "__main__":
    mcp = MCPServer()
    mcp.start_scheduler()
