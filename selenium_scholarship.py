import json
import re
from typing import List, Dict, Any
from dataclasses import dataclass
from enum import Enum
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from datetime import datetime

class Level(Enum):
    BACHELOR = "學士"
    MASTER = "碩士"
    DOCTOR = "博士"

class Identity(Enum):
    LOCAL = "本國人"
    OVERSEAS_CHINESE = "僑生"
    INTERNATIONAL = "外籍生"

class StudyType(Enum):
    FULL_TIME = "全日"
    PART_TIME = "在職專班"

class Source(Enum):
    OVERSEAS_AFFAIRS = "僑陸組"
    STUDENT_AFFAIRS = "生輔組"
    CSIE = "資工系"

@dataclass
class UserInput:
    department: str
    level: Level
    year: int
    identity: Identity
    study_type: StudyType

@dataclass
class Scholarship:
    title: str
    url: str
    source: str
    date: str = ""
    description: str = ""
    status: str = ""
    category: str = ""
    deadline: str = ""
    amount: str = ""
    contact: str = ""

class ScholarshipCrawler:
    def __init__(self):
        # URLs específicas actualizadas
        self.target_urls = {
            "生輔組": "https://advisory.ntu.edu.tw/CMS/Scholarship?pageId=232",
            "資工系": "https://www.csie.ntu.edu.tw/zh_tw/Announcements/11", 
            "僑陸組": "https://gocfs.ntu.edu.tw/board/index/tab/1"
        }
        
        self.departments = {
            "資訊工程學系": "資工系",
            "電機工程學系": "電機系", 
            "機械工程學系": "機械系",
            "化學工程學系": "化工系",
            "土木工程學系": "土木系",
            "資工系": "資工系",
            "電機系": "電機系",
            "機械系": "機械系",
            "化工系": "化工系",
            "土木系": "土木系"
        }
    
    def setup_driver(self):
        """設置 Chrome WebDriver"""
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--ignore-ssl-errors")
        options.add_argument("--disable-web-security")
        options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    def crawl_student_affairs(self, max_pages=3) -> List[Scholarship]:
        """爬取生輔組獎學金 """
        scholarships = []
        driver = self.setup_driver()
        
        try:
            url = self.target_urls["生輔組"]
            print(f"正在爬取生輔組: {url}")
            driver.get(url)
            time.sleep(5)  # 增加等待時間
            
            # 等待頁面加載
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 嘗試多種選擇器來找獎學金項目
            selectors = [
                "div.scholarship-list .item",
                "div.list-group .list-group-item",
                "table.table tbody tr",
                "div.row .col-md-12",
                "div[class*='scholarship']",
                ".news-item",
                "li.list-group-item",
                "div.panel div.panel-body"
            ]
            
            items = []
            for selector in selectors:
                try:
                    items = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(items) > 0:
                        print(f"找到 {len(items)} 個項目 (使用選擇器: {selector})")
                        break
                except Exception as e:
                    continue
            
            # 如果還是沒找到，嘗試更通用的方法
            if not items:
                items = driver.find_elements(By.XPATH, "//a[contains(text(), '獎學金') or contains(text(), '獎助')]")
            
            for item in items:
                scholarship = self.parse_scholarship_item(item, "生輔組", driver)
                if scholarship:
                    scholarships.append(scholarship)
                    
        except Exception as e:
            print(f"爬取生輔組時出錯: {e}")
        finally:
            driver.quit()
            
        return scholarships
    
    def crawl_csie(self, max_pages=3) -> List[Scholarship]:
        """爬取資工系 """
        scholarships = []
        driver = self.setup_driver()
        
        try:
            url = self.target_urls["資工系"]
            print(f"正在爬取資工系: {url}")
            driver.get(url)
            time.sleep(5)
            
            # 等待頁面加載
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 資工系可能的選擇器
            selectors = [
                "div.announcement-list .item",
                "table.table tbody tr",
                "div.news-list .news-item",
                "ul.list-group li",
                "div[class*='announcement']",
                "div[class*='news']",
                ".content-list .item",
                "tr"
            ]
            
            items = []
            for selector in selectors:
                try:
                    items = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(items) > 1:  # 至少要有2個以上才算找到列表
                        print(f"找到 {len(items)} 個項目 (使用選擇器: {selector})")
                        break
                except Exception as e:
                    continue
            
            # 如果還是沒找到，嘗試找所有連結
            if not items:
                items = driver.find_elements(By.TAG_NAME, "a")
            
            for item in items:
                scholarship = self.parse_scholarship_item(item, "資工系", driver)
                if scholarship:
                    scholarships.append(scholarship)
                    
        except Exception as e:
            print(f"爬取資工系時出錯: {e}")
        finally:
            driver.quit()
            
        return scholarships
    
    def crawl_overseas_affairs(self, max_pages=3) -> List[Scholarship]:
        """爬取僑陸組 """
        scholarships = []
        driver = self.setup_driver()
        
        try:
            url = self.target_urls["僑陸組"]
            print(f"正在爬取僑陸組: {url}")
            driver.get(url)
            time.sleep(5)
            
            # 等待頁面加載
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 僑陸組可能的選擇器
            selectors = [
                "div.content-list .item",
                "table tbody tr",
                "ul li",
                "div[class*='list'] .item",
                ".news-list .news-item",
                "div.row div[class*='col']",
                "a[href*='scholarship']",
                "tr"
            ]
            
            items = []
            for selector in selectors:
                try:
                    items = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(items) > 0:
                        print(f"找到 {len(items)} 個項目 (使用選擇器: {selector})")
                        break
                except Exception as e:
                    continue
            
            for item in items:
                scholarship = self.parse_scholarship_item(item, "僑陸組", driver)
                if scholarship:
                    scholarships.append(scholarship)
                    
        except Exception as e:
            print(f"爬取僑陸組時出錯: {e}")
        finally:
            driver.quit()
            
        return scholarships
    
    def parse_scholarship_item(self, item, source_name, driver):
        """解析單個獎學金項目"""
        try:
            # 嘗試找到連結和標題
            a_tag = None
            title = ""
            href = ""
            
            # 如果item本身就是a標籤
            if item.tag_name == 'a':
                a_tag = item
                title = item.text.strip()
                href = item.get_attribute("href")
            else:
                # 嘗試在item內找a標籤
                try:
                    a_tag = item.find_element(By.TAG_NAME, "a")
                    title = a_tag.text.strip()
                    href = a_tag.get_attribute("href")
                except:
                    # 如果沒有a標籤，嘗試獲取文本
                    title = item.text.strip()
                    href = ""
            
            # 檢查是否為獎學金相關
            if not title or not self.is_scholarship_related(title):
                return None
            
            # 處理相對連結
            if href and not href.startswith("http"):
                if source_name == "生輔組":
                    base_url = "https://advisory.ntu.edu.tw"
                elif source_name == "資工系":
                    base_url = "https://www.csie.ntu.edu.tw"
                elif source_name == "僑陸組":
                    base_url = "https://gocfs.ntu.edu.tw"
                else:
                    base_url = ""
                
                if base_url:
                    href = base_url + ("" if href.startswith("/") else "/") + href
            
            # 提取其他信息
            date = self.extract_date_from_element(item)
            status = self.extract_status_from_element(item)
            category = self.extract_category_from_element(item)
            
            scholarship = Scholarship(
                title=title,
                url=href or "",
                source=source_name,
                date=date,
                status=status,
                category=category
            )
            
            return scholarship
            
        except Exception as e:
            return None
    
    def extract_date_from_element(self, element):
        """從元素中提取日期"""
        try:
            # 先嘗試找特定的日期元素
            date_selectors = [
                "span.date",
                ".time",
                "[class*='date']",
                "td:last-child",  # 表格最後一欄通常是日期
                "small"
            ]
            
            for selector in date_selectors:
                try:
                    date_elem = element.find_element(By.CSS_SELECTOR, selector)
                    date_text = date_elem.text.strip()
                    if self.is_date_format(date_text):
                        return date_text
                except:
                    continue
            
            # 如果沒找到，嘗試從文本中提取
            text = element.text
            date_patterns = [
                r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})',
                r'(\d{1,2}[-/]\d{1,2}[-/]\d{4})',
                r'(\d{4}年\d{1,2}月\d{1,2}日)',
                r'(\d{4}\.\d{1,2}\.\d{1,2})'
            ]
            
            for pattern in date_patterns:
                match = re.search(pattern, text)
                if match:
                    return match.group(1)
                    
        except:
            pass
        
        return ""
    
    def is_date_format(self, text):
        """檢查文字是否為日期格式"""
        date_patterns = [
            r'\d{4}[-/]\d{1,2}[-/]\d{1,2}',
            r'\d{1,2}[-/]\d{1,2}[-/]\d{4}',
            r'\d{4}年\d{1,2}月\d{1,2}日',
            r'\d{4}\.\d{1,2}\.\d{1,2}'
        ]
        
        for pattern in date_patterns:
            if re.search(pattern, text):
                return True
        return False
    
    def extract_status_from_element(self, element):
        """從元素中提取狀態"""
        try:
            text = element.text.lower()
            status_keywords = {
                '開放申請': ['開放', '申請中', '受理中'],
                '截止': ['截止', '結束', 'closed'],
                '審核中': ['審核', '評選'],
                '即將截止': ['即將', 'deadline']
            }
            
            for status, keywords in status_keywords.items():
                if any(keyword in text for keyword in keywords):
                    return status
        except:
            pass
        
        return ""
    
    def extract_category_from_element(self, element):
        """從元素中提取類別"""
        try:
            text = element.text
            categories = []
            
            category_keywords = {
                '研究生': ['研究生', '碩士', '博士', 'graduate'],
                '大學生': ['大學生', '學士', 'undergraduate'],
                '清寒': ['清寒', '低收入'],
                '優秀': ['優秀', '績優'],
                '僑生': ['僑生'],
                '外籍生': ['外籍', 'international']
            }
            
            for category, keywords in category_keywords.items():
                if any(keyword in text for keyword in keywords):
                    categories.append(category)
            
            return ', '.join(categories)
        except:
            pass
        
        return ""
    
    def is_scholarship_related(self, title: str) -> bool:
        """判斷標題是否與獎學金相關"""
        if not title or len(title.strip()) < 3:
            return False
            
        scholarship_keywords = [
            "獎學金", "scholarship", "獎助", "補助", "津貼", 
            "助學", "獎勵", "獎助學金", "教育基金", "學費減免"
        ]
        
        title_lower = title.lower()
        return any(keyword in title_lower for keyword in scholarship_keywords)
    
    def filter_scholarships(self, scholarships: List[Scholarship], user_input: UserInput) -> List[Scholarship]:
        """根據用戶條件過濾獎學金"""
        filtered = []
        
        for scholarship in scholarships:
            title_lower = scholarship.title.lower()
            is_match = False
            
            # 根據身份篩選
            if user_input.identity == Identity.OVERSEAS_CHINESE:
                if "僑生" in title_lower or "overseas" in title_lower or scholarship.source == "僑陸組":
                    is_match = True
            elif user_input.identity == Identity.INTERNATIONAL:
                if "外籍" in title_lower or "international" in title_lower:
                    is_match = True
            else:  # 本國人
                if "僑生" not in title_lower and "外籍" not in title_lower:
                    is_match = True
            
            # 根據學位層級篩選
            level_keywords = {
                Level.BACHELOR: ["學士", "大學", "undergraduate"],
                Level.MASTER: ["碩士", "研究所", "master", "graduate"],
                Level.DOCTOR: ["博士", "phd", "doctoral"]
            }
            
            if any(keyword in title_lower for keyword in level_keywords.get(user_input.level, [])):
                is_match = True
            
            # 根據系所篩選
            if user_input.department in title_lower or scholarship.source.endswith(user_input.department):
                is_match = True
            
            if is_match:
                filtered.append(scholarship)
        
        # 如果沒有嚴格匹配的結果，放寬條件
        if not filtered:
            filtered = [s for s in scholarships if self.is_potentially_relevant(s, user_input)]
        
        return filtered
    
    def is_potentially_relevant(self, scholarship: Scholarship, user_input: UserInput) -> bool:
        """寬鬆條件判斷獎學金是否可能相關"""
        if (user_input.identity == Identity.OVERSEAS_CHINESE and scholarship.source == "僑陸組") or \
           (scholarship.source == "生輔組") or \
           (scholarship.source.endswith(user_input.department)):
            return True
            
        return False

class ScholarshipFinder:
    def __init__(self):
        self.crawler = ScholarshipCrawler()
        
    def get_available_departments(self) -> List[str]:
        """獲取可用的系所列表"""
        return list(self.crawler.departments.keys())
    
    def search_scholarships(self, user_input: UserInput, max_pages_per_source=3) -> Dict[str, Any]:
        """搜尋獎學金"""
        print("開始搜尋獎學金...")
        
        all_scholarships = []
        
        # 根據身份決定爬取來源
        if user_input.identity == Identity.OVERSEAS_CHINESE:
            print("正在爬取僑陸組...")
            all_scholarships.extend(self.crawler.crawl_overseas_affairs(max_pages_per_source))
        
        # 所有人都可能適用生輔組獎學金
        print("正在爬取生輔組...")
        all_scholarships.extend(self.crawler.crawl_student_affairs(max_pages_per_source))
        
        # 如果是資工系相關
        if "資工" in user_input.department:
            print("正在爬取資工系...")
            all_scholarships.extend(self.crawler.crawl_csie(max_pages_per_source))
        
        print(f"共爬取到 {len(all_scholarships)} 個獎學金項目")
        
        # 根據用戶條件過濾
        filtered_scholarships = self.crawler.filter_scholarships(all_scholarships, user_input)
        print(f"過濾後剩餘 {len(filtered_scholarships)} 個相關項目")
        
        return {"data": filtered_scholarships}
    
    def save_to_excel(self, scholarships: List[Scholarship], filename=None):
        """將獎學金數據保存到Excel文件"""
        if not filename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'scholarships_{timestamp}.xlsx'
        
        # 轉換為字典列表
        data = []
        for scholarship in scholarships:
            data.append({
                'title': scholarship.title,
                'source': scholarship.source,
                'date': scholarship.date,
                'status': scholarship.status,
                'category': scholarship.category,
                'url': scholarship.url,
                'description': scholarship.description,
                'deadline': scholarship.deadline,
                'amount': scholarship.amount,
                'contact': scholarship.contact,
                'crawl_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
        
        # 創建DataFrame
        df = pd.DataFrame(data)
        
        if df.empty:
            print("沒有數據可以保存")
            return None
        
        # 保存到Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='獎學金列表', index=False)
            
            # 調整欄位寬度
            worksheet = writer.sheets['獎學金列表']
            for column in worksheet.columns:
                max_length = 0
                column_name = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_name].width = adjusted_width
        
        print(f"數據已保存到: {filename}")
        return filename

def main():
    finder = ScholarshipFinder()
    
    print("=== 獎學金查詢系統 - 特定URL版本 ===")
    print("支援網站:")
    print("• 生輔組: https://advisory.ntu.edu.tw/CMS/Scholarship?pageId=232")
    print("• 資工系: https://www.csie.ntu.edu.tw/zh_tw/Announcements/11")  
    print("• 僑陸組: https://gocfs.ntu.edu.tw/page/index/menu_sn/61")
    print("可用系所:", ", ".join(finder.get_available_departments()))
    
    # 示例輸入
    user_input_data = {
        "department": "資工系",
        "level": "碩士",
        "year": 2,
        "identity": "僑生",
        "study_type": "全日"
    }
    
    try:
        # 轉換輸入格式
        user_input = UserInput(
            department=user_input_data["department"],
            level=Level.MASTER if user_input_data["level"] == "碩士" else Level.BACHELOR,
            year=user_input_data["year"],
            identity=Identity.OVERSEAS_CHINESE if user_input_data["identity"] == "僑生" else Identity.LOCAL,
            study_type=StudyType.FULL_TIME if user_input_data["study_type"] == "全日" else StudyType.PART_TIME
        )
        
        print(f"\n查詢條件: {user_input}")
        
        # 搜尋獎學金
        result = finder.search_scholarships(user_input, max_pages_per_source=3)
        
        # 輸出JSON結果
        data_for_json = []
        for scholarship in result["data"]:
            data_for_json.append({
                "title": scholarship.title,
                "url": scholarship.url,
                "source": scholarship.source,
                "date": scholarship.date,
                "status": scholarship.status,
                "category": scholarship.category
            })
        
        print(f"\n=== JSON 查詢結果 ===")
        print(json.dumps({"data": data_for_json}, ensure_ascii=False, indent=2))
        
        # 保存到Excel
        if result["data"]:
            excel_file = finder.save_to_excel(result["data"])
            print(f"\n✅ Excel 文件已保存: {excel_file}")
        else:
            print("\n沒有找到相關獎學金")
        
    except Exception as e:
        print(f"查詢過程中出錯: {e}")

if __name__ == "__main__":
    main()