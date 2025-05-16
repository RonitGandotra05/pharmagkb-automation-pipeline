import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import logging
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
import os

def setup_logging():
    """Set up logging with detailed format"""
    log_dir = Path('test_output/logs')
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = log_dir / f'test_processing_{timestamp}.log'
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return log_file

class TestPharmGKBScraper:
    def __init__(self):
        self.setup_logging()
        self.setup_output_dirs()
        
    def setup_logging(self):
        log_dir = os.path.join('logs')
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = os.path.join(log_dir, f'scraping_{timestamp}.log')
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
    
    def setup_output_dirs(self):
        """Create necessary output directories"""
        self.scraped_content_dir = os.path.join('step2_input', 'scraped_content')
        os.makedirs(self.scraped_content_dir, exist_ok=True)
        logging.info(f"Created output directory: {self.scraped_content_dir}")

    def wait_for_data_load(self, driver):
        """Wait for page content to load"""
        logging.info("Waiting for page content to load...")
        wait = WebDriverWait(driver, 45)
        
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
            time.sleep(20)  # Wait for dynamic content
            content = driver.find_element(By.TAG_NAME, 'body').text
            return len(content.strip()) > 0
        except Exception as e:
            logging.error(f"Error waiting for data load: {str(e)}")
            return False

    def save_content(self, sample_id, content):
        """Save scraped content to file"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = os.path.join(self.scraped_content_dir, f"{sample_id}_processed_content_{timestamp}.txt")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
            
        logging.info(f"Saved content to: {output_file}")
        return output_file

    def process_content(self, soup):
        """Process and clean content"""
        logging.info("Processing content...")
        text_content = soup.get_text(separator='\n', strip=True)
        
        # Find the start marker
        start_marker = "Learn more about CPIC"
        start_idx = text_content.find(start_marker)
        
        if start_idx != -1:
            return text_content[start_idx:]
        return text_content

    def scrape_test_sample(self):
        """Scrape data for the test sample"""
        logging.info(f"Starting test scrape for sample: {self.sample_id}")
        
        try:
            # Read URL from CSV
            df = pd.read_csv('consolidated_sample_urls_JAPAN(in).csv')
            sample_row = df[df['Sample'] == self.sample_id]
            
            if sample_row.empty:
                logging.error(f"No URL found for sample {self.sample_id}")
                return None
                
            url = sample_row.iloc[0]['Generated URL']
            logging.info(f"Found URL for sample {self.sample_id}: {url}")
            
            print(f"\nProcessing test sample: {self.sample_id}")
            print(f"URL: {url}\n")
            
            driver = None
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=self.options)
                driver.get(url)
                
                if not self.wait_for_data_load(driver):
                    logging.error("Failed to load page content")
                    return None

                # Process and save content
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                processed_content = self.process_content(soup)
                content_path = self.save_content(self.sample_id, processed_content)
                
                if content_path:
                    logging.info(f"Successfully processed test sample {self.sample_id}")
                    print(f"Content saved to: {content_path}")
                    return content_path
                    
            except Exception as e:
                logging.error(f"Error processing test sample: {str(e)}")
                return None
                
            finally:
                if driver:
                    driver.quit()
                    
        except Exception as e:
            logging.error(f"Error in scrape_test_sample: {str(e)}")
            return None

def main():
    """Main function to run test scraper"""
    log_file = setup_logging()
    logging.info("Starting test scraper process...")
    print(f"Log file: {log_file}")
    
    # Get list of all files from step1 output directory
    step1_dir = Path('step1_phenotype_genotype_output/processed_files')
    if not step1_dir.exists():
        logging.error(f"Directory not found: {step1_dir}")
        return
        
    # Process each file in the directory
    total_files = len(list(step1_dir.glob('*.xlsx')))
    processed_count = 0
    failed_count = 0
    
    print(f"\nFound {total_files} files to process")
    
    for xlsx_file in step1_dir.glob('*.xlsx'):
        sample_id = xlsx_file.stem  # Get filename without extension
        processed_count += 1
        
        print(f"\nProcessing file {processed_count}/{total_files}: {sample_id}")
        logging.info(f"Processing sample: {sample_id}")
        
        try:
            scraper = TestPharmGKBScraper(sample_id)
            content_path = scraper.scrape_test_sample()
            
            if content_path:
                logging.info(f"Successfully processed sample {sample_id}")
                print(f"Content saved to: {content_path}")
            else:
                failed_count += 1
                logging.error(f"Failed to process sample {sample_id}")
                print(f"Failed to process sample: {sample_id}")
                
        except Exception as e:
            failed_count += 1
            logging.error(f"Error processing sample {sample_id}: {str(e)}")
            print(f"Error processing sample {sample_id}: {str(e)}")
            continue
            
    # Print summary
    print(f"\nProcessing completed!")
    print(f"Total files processed: {processed_count}")
    print(f"Successfully processed: {processed_count - failed_count}")
    print(f"Failed to process: {failed_count}")
    logging.info(f"Processing completed. Success: {processed_count - failed_count}, Failed: {failed_count}")

if __name__ == "__main__":
    main() 