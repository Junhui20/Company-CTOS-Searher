import time
import re
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

logger = logging.getLogger(__name__)

CTOS_BASE_URL = "https://businessreport.ctoscredit.com.my/oneoffreport/home"


class CTOSScraper:
    def __init__(self, headless=False, fast_mode=True):
        self.driver = None
        self.headless = headless
        self.fast_mode = fast_mode
        self.base_url = CTOS_BASE_URL

    def start_driver(self):
        """Initializes the Chrome connection."""
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--window-size=1280,800")
        options.add_argument("--log-level=3")

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=options
        )
        self.driver.implicitly_wait(5)
        logger.info("Browser started.")

    def close_driver(self):
        if self.driver:
            self.driver.quit()
            logger.info("Browser closed.")

    @staticmethod
    def _without_element(data):
        """Returns a new dict without the 'element' key (avoids mutating the original)."""
        return {k: v for k, v in data.items() if k != "element"}

    @staticmethod
    def _results_to_serializable(results):
        """Returns a new list of dicts without selenium element references."""
        return [{k: v for k, v in r.items() if k != "element"} for r in results]

    def search_company(self, company_name):
        """
        Main entry point for searching a company.
        Returns a dictionary:
        {
            "status": "FOUND" | "NOT_FOUND" | "AMBIGUOUS" | "ERROR",
            "data": { ... } or [list of candidates],
            "message": "..."
        }
        """
        if not self.driver:
            self.start_driver()

        clean_name = company_name.strip()

        logger.info("Searching for: %s", clean_name)
        try:
            result = self._perform_search(clean_name)
        except Exception as e:
            logger.warning("Driver crashed? Restarting. Error: %s", e)
            self.close_driver()
            self.start_driver()
            result = self._perform_search(clean_name)

        if result["status"] in ["FOUND", "FOUND_PARTIAL"]:
            return result

        if result["status"] == "AMBIGUOUS":
            companies_only = [
                c for c in result["data"] if "COMPANY" in c.get("type", "").upper()
            ]
            logger.info(
                "Ambiguity check: Found %d total. Filtered 'Company': %d",
                len(result["data"]),
                len(companies_only),
            )

            if len(companies_only) == 1:
                logger.info("Auto-resolved ambiguity by filtering for 'Company'.")
                if self.fast_mode:
                    logger.info("Fast Mode: Returning resolved company without details.")
                    return {
                        "status": "FOUND",
                        "data": self._without_element(companies_only[0]),
                    }
                return self._click_and_scrape_details(companies_only[0])
            elif len(companies_only) > 1:
                return {**result, "data": companies_only}
            return result

        # Strategy 2: Fuzzy / Cleaned Name
        reduced_name = self._clean_company_name(clean_name)
        if reduced_name and reduced_name.lower() != clean_name.lower():
            logger.info("Retry with reduced name: %s", reduced_name)
            try:
                result_fuzzy = self._perform_search(reduced_name)
                if result_fuzzy["status"] not in ["NOT_FOUND", "ERROR"]:
                    return result_fuzzy
            except Exception as e:
                logger.warning("Fuzzy retry failed: %s", e)

        return {
            "status": "NOT_FOUND",
            "data": [],
            "message": "No results found after retries.",
        }

    def _clean_company_name(self, name):
        """Removes common suffixes and bracketed text to improve search hit rate."""
        name = re.sub(r"\(.*?\)", "", name)
        name = re.sub(r"\s+SDN\.?\s*BHD\.?", "", name, flags=re.IGNORECASE)
        name = re.sub(r"\s+BHD\.?", "", name, flags=re.IGNORECASE)
        name = re.sub(r"\s+PLT", "", name, flags=re.IGNORECASE)
        name = re.sub(r"\s+ENTERPRISE", "", name, flags=re.IGNORECASE)
        return name.strip()

    def _perform_search(self, term):
        try:
            self.driver.get(self.base_url)

            search_box = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "searchKey"))
            )
            search_box.clear()
            search_box.send_keys(term)

            try:
                btn = self.driver.find_element(By.CSS_SELECTOR, ".search_txt_home")
            except NoSuchElementException:
                btn = self.driver.find_element(By.CSS_SELECTOR, ".search_txt")
            btn.click()

            WebDriverWait(self.driver, 10).until(
                lambda d: d.find_elements(By.CSS_SELECTOR, ".mat-row")
                or "No result found" in d.page_source
                or "0 result" in d.page_source
            )

            page_src = self.driver.page_source
            if "No result found" in page_src or "0 result" in page_src:
                return {"status": "NOT_FOUND", "data": []}

            rows = self.driver.find_elements(By.CSS_SELECTOR, ".mat-row")
            results = []

            for row in rows:
                try:
                    cols = row.find_elements(By.CSS_SELECTOR, ".mat-cell")
                    try:
                        logger.debug("Row cols: %s", [c.text for c in cols])
                    except Exception:
                        pass

                    if len(cols) >= 3:
                        reg_no = cols[0].text.strip()
                        comp_name = cols[1].text.strip()
                        comp_type = cols[2].text.strip()

                        results.append(
                            {
                                "reg_no": reg_no,
                                "name": comp_name,
                                "type": comp_type,
                                "element": row,
                            }
                        )
                except Exception as e:
                    logger.warning("Error parsing row: %s", e)

            if not results:
                return {"status": "NOT_FOUND", "data": []}

            if len(results) == 1:
                if self.fast_mode:
                    logger.info(
                        "Fast Mode: Found 1 result for '%s'. Returning summary.", term
                    )
                    return {
                        "status": "FOUND",
                        "data": self._without_element(results[0]),
                    }

                click_result = self._click_and_scrape_details(results[0])
                if click_result["status"] in ["FOUND", "FOUND_PARTIAL"]:
                    return click_result
                return {
                    "status": "FOUND_PARTIAL",
                    "data": self._without_element(results[0]),
                }

            return {
                "status": "AMBIGUOUS",
                "data": self._results_to_serializable(results),
            }

        except Exception as e:
            logger.error("Search error for '%s': %s", term, e)
            return {"status": "ERROR", "message": str(e)}

    def _click_and_scrape_details(self, result_item):
        """Clicks a result in the list and scrapes the detail page."""
        try:
            target_reg = result_item["reg_no"]
            rows = self.driver.find_elements(By.CSS_SELECTOR, ".mat-row")
            target_row = None
            for r in rows:
                if target_reg in r.text:
                    target_row = r
                    break

            if not target_row:
                return {
                    "status": "FOUND_PARTIAL",
                    "data": self._without_element(result_item),
                }

            try:
                link = target_row.find_element(
                    By.CSS_SELECTOR, ".mat-column-Company_Name"
                )

                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", link
                )
                time.sleep(0.5)

                try:
                    link.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", link)

                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "//div[contains(text(), 'Date of Registration')]",
                        )
                    )
                )

                details = {
                    "reg_no": result_item["reg_no"],
                    "name": result_item["name"],
                    "type": result_item["type"],
                }

                def get_text_by_label(label):
                    try:
                        xpath = (
                            f"//div[contains(@class,'label') and contains(text(), '{label}')]"
                            "/following-sibling::div"
                        )
                        el = self.driver.find_element(By.XPATH, xpath)
                        return el.text.strip()
                    except NoSuchElementException:
                        return "-"

                details["nature_of_business"] = get_text_by_label("Nature of Business")
                details["date_of_registration"] = get_text_by_label(
                    "Date of Registration"
                )
                details["state"] = get_text_by_label("State")

                return {"status": "FOUND", "data": details}

            except Exception as e:
                logger.info("Detail page not reached: %s", e)
                return {
                    "status": "FOUND_PARTIAL",
                    "data": self._without_element(result_item),
                }

        except Exception as e:
            logger.warning(
                "Detail Scraping Error (fallback): %s", str(e).splitlines()[0]
            )
            return {
                "status": "FOUND_PARTIAL",
                "data": self._without_element(result_item),
            }


# Test block
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    bot = CTOSScraper(headless=False)
    # res = bot.search_company("Ace Altair Travels Sdn. Bhd")
    # print(res)
    # bot.close_driver()
