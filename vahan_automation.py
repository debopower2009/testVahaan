from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import glob
import pandas as pd

class GitHubActionsVahanProcessor:
    def __init__(self):
        # Use current working directory for downloads in GitHub Actions
        self.download_path = os.path.join(os.getcwd(), 'downloads')
        os.makedirs(self.download_path, exist_ok=True)
        self.driver = None
        
    def setup_headless_chrome(self):
        """Setup Chrome for GitHub Actions (headless mode)"""
        chrome_options = Options()
        
        # Essential for GitHub Actions
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        
        # Download preferences
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        return self.driver
    
    def wait_for_download(self, timeout=120):
        """Wait for file download completion"""
        print("🔍 Waiting for download...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            files = glob.glob(os.path.join(self.download_path, '*.xlsx'))
            
            if files:
                # Get the most recent file
                latest_file = max(files, key=os.path.getctime)
                
                # Check if file is stable (not being written)
                initial_size = os.path.getsize(latest_file)
                time.sleep(3)
                final_size = os.path.getsize(latest_file)
                
                if initial_size == final_size and final_size > 1000:
                    print(f"✅ Download completed: {os.path.basename(latest_file)}")
                    return latest_file
            
            time.sleep(2)
        
        return None
    
    def select_dropdown(self, dropdown_id, option_text, wait_time=15):
        """Select dropdown option"""
        wait = WebDriverWait(self.driver, wait_time)
        
        try:
            # Click dropdown
            dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
            dropdown.click()
            time.sleep(2)
            
            # Select option
            panel_id = dropdown_id.replace("_label", "_panel")
            option_xpath = f"//div[@id='{panel_id}']//li[@data-label='{option_text}']"
            option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
            option.click()
            time.sleep(2)
            
            print(f"✅ Selected: {option_text}")
            return True
            
        except Exception as e:
            print(f"❌ Failed to select {option_text}: {e}")
            return False
    
    def run_automation(self):
        """Run the complete automation"""
        try:
            print("🚀 Starting GitHub Actions Vahan Automation")
            
            # Setup Chrome
            self.setup_headless_chrome()
            wait = WebDriverWait(self.driver, 30)
            
            # Navigate to Vahan dashboard
            url = "https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml"
            print(f"🌐 Opening: {url}")
            self.driver.get(url)
            
            # Wait for page load
            wait.until(EC.presence_of_element_located((By.ID, "yaxisVar_label")))
            print("✅ Page loaded")
            
            # Configure parameters
            print("📊 Configuring parameters...")
            
            # Y-axis: Maker
            if not self.select_dropdown("yaxisVar_label", "Maker"):
                return None
            time.sleep(3)
            
            # X-axis: Fuel  
            if not self.select_dropdown("xaxisVar_label", "Fuel"):
                return None
            time.sleep(3)
            
            # Refresh
            print("🔄 Refreshing...")
            refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, "j_idt75")))
            refresh_btn.click()
            time.sleep(8)
            
            # Select month
            print("📅 Selecting SEP...")
            try:
                wait.until(EC.presence_of_element_located((By.ID, "groupingTable:selectMonth_label")))
                if not self.select_dropdown("groupingTable:selectMonth_label", "SEP"):
                    # Fallback method
                    month_input = self.driver.find_element(By.ID, "groupingTable:selectMonth_input")
                    self.driver.execute_script("arguments[0].value = 'SEP';", month_input)
                    self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", month_input)
            except Exception as e:
                print(f"Month selection error: {e}")
            
            time.sleep(3)
            
            # Download
            print("⬇️ Starting download...")
            download_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:j_idt91")))
            download_btn.click()
            
            # Wait for download
            downloaded_file = self.wait_for_download()
            
            if downloaded_file:
                print(f"🎉 Downloaded: {os.path.basename(downloaded_file)}")
                return self.process_file(downloaded_file)
            else:
                print("❌ Download failed")
                return None
                
        except Exception as e:
            print(f"❌ Automation failed: {e}")
            return None
        
        finally:
            if self.driver:
                self.driver.quit()
    
    def process_file(self, file_path):
        """Process the downloaded Excel file"""
        try:
            print(f"📈 Processing: {os.path.basename(file_path)}")
            
            # Load and process data (your existing logic)
            df = pd.read_excel(file_path, sheet_name='reportTable', skiprows=3)
            df = df.drop(df.columns[0], axis=1)
            df = df.rename(columns={df.columns[0]: 'Makers'})
            
            # Filter makers
            makers_to_filter = [
                "ATHER ENERGY LTD",
                "BAJAJ AUTO LTD",
                "HERO ELECTRIC VEHICLE PVT LTD",
                "HERO ELECTRIC VEHICLES PVT. LTD",
                "HERO HONDA MOTORS  LTD",
                "HERO MOTOCORP LTD",
                "OLA ELECTRIC TECHNOLOGIES PVT LTD",
                "TVS MOTOR COMPANY LTD"
            ]
            
            filtered_df = df[df['Makers'].isin(makers_to_filter)].copy()
            
            # Group Hero companies
            hero_makers = [
                "HERO ELECTRIC VEHICLE PVT LTD",
                "HERO ELECTRIC VEHICLES PVT. LTD",
                "HERO HONDA MOTORS  LTD",
                "HERO MOTOCORP LTD"
            ]
            filtered_df.loc[filtered_df['Makers'].isin(hero_makers), 'Makers'] = 'HERO ELECTRIC'
            
            # Process numeric columns
            numeric_columns = ['ELECTRIC(BOV)', 'PLUG-IN HYBRID EV', 'PURE EV', 'STRONG HYBRID EV']
            
            for col in numeric_columns:
                if col in filtered_df.columns:
                    filtered_df[col] = filtered_df[col].astype(str).str.replace(",", "")
                    filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce')
            
            # Group and calculate totals
            result_df = filtered_df.groupby('Makers', as_index=False).sum()
            available_columns = [col for col in numeric_columns if col in result_df.columns]
            selected_df = result_df.loc[:, ['Makers'] + available_columns]
            selected_df['TOTAL EV'] = selected_df[available_columns].sum(axis=1)
            
            # Sort by total EV
            sorted_df = selected_df.sort_values(by='TOTAL EV', ascending=False)
            
            # Save processed results
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_file = f"vahan_ev_report_{timestamp}.xlsx"
            sorted_df.to_excel(output_file, index=False)
            
            print(f"💾 Saved: {output_file}")
            print("\n🎉 RESULTS:")
            print(sorted_df)
            
            return sorted_df
            
        except Exception as e:
            print(f"❌ Processing failed: {e}")
            return None

# Main execution
if __name__ == "__main__":
    processor = GitHubActionsVahanProcessor()
    results = processor.run_automation()
    
    if results is not None:
        print(f"\n✅ SUCCESS: Processed {len(results)} companies")
    else:
        print("\n❌ FAILED: Automation unsuccessful")
