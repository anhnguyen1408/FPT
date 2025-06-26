import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import os

def chuan_hoa_text(text):
    if isinstance(text, str):
        return text.strip().replace('​', '')
    return str(text).strip()

def doc_du_lieu(file_path):
    df = pd.read_excel(file_path, dtype=str)
    df = df[['Mã số thuế', 'Mã tra cứu', 'URL']]
    df = df.applymap(chuan_hoa_text)
    return df

def tao_trinh_duyet():
    download_path = os.path.abspath("hoa_don")
    os.makedirs(download_path, exist_ok=True)

    options = Options()
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def tra_cuu_fpt(driver, url, mst, ma_tra_cuu):
    driver.get(url)
    try:
        wait = WebDriverWait(driver, 15)

        input_mst = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//input[@placeholder="MST bên bán"]')
        ))
        input_mst.clear()
        input_mst.send_keys(mst)

        input_matc = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//input[@placeholder="Mã tra cứu hóa đơn"]')
        ))
        input_matc.clear()
        input_matc.send_keys(ma_tra_cuu)
        input_matc.send_keys(Keys.ENTER)

        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))
        driver.switch_to.frame(iframe)

        try:
            first_row = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'table tbody tr')))
            first_row.click()
            time.sleep(2)
        except:
            driver.switch_to.default_content()
            return "Không có kết quả để tải"

        try:
            btn_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[span[contains(text(), "PDF")]]')))
            btn_pdf.click()
        except:
            pass

        try:
            btn_xml = wait.until(EC.element_to_be_clickable((
                By.XPATH, '//button[contains(@class, "webix_button") and contains(@class, "webix_img_btn")]'
            )))
            btn_xml.click()
        except:
            pass

        time.sleep(4)
        driver.switch_to.default_content()
        return "Đã tải PDF/XML nếu có"

    except Exception as e:
        driver.switch_to.default_content()
        return f"Lỗi khi tra cứu: {str(e)}"

def tra_cuu_misa(driver, url, ma_tra_cuu):
    driver.get(url)
    try:
        wait = WebDriverWait(driver, 10)
        input_code = wait.until(EC.presence_of_element_located((By.ID, "txtCode")))
        input_code.clear()
        input_code.send_keys(ma_tra_cuu)

        btn_search = wait.until(EC.element_to_be_clickable((By.ID, "btnSearchInvoice")))
        btn_search.click()

        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'm-invoice-info')))
        return "Thành công"
    except Exception as e:
        return f"Lỗi: {str(e)}"

def tra_cuu_ehoadon(driver, url, ma_tra_cuu):
    driver.get(url + ma_tra_cuu)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "invoice-info")))
        return "Thành công"
    except Exception as e:
        return f"Lỗi: {str(e)}"

def chay_tra_cuu():
    df = doc_du_lieu("input.xlsx")

    ket_qua = []
    driver = tao_trinh_duyet()

    for index, row in df.iterrows():
        mst = row['Mã số thuế'].strip()
        ma_tra_cuu = row['Mã tra cứu'].strip()
        url = row['URL'].strip().lower()

        if "fpt" in url:
            trang_thai = tra_cuu_fpt(driver, url, mst, ma_tra_cuu)
        elif "meinvoice.vn" in url:
            trang_thai = tra_cuu_misa(driver, url, ma_tra_cuu)
        elif "van.ehoadon.vn" in url:
            trang_thai = tra_cuu_ehoadon(driver, url, ma_tra_cuu)
        else:
            trang_thai = "Không hỗ trợ URL này"

        ket_qua.append({
            "Mã số thuế": mst,
            "Mã tra cứu": ma_tra_cuu,
            "URL": url,
            "Trạng thái": trang_thai
        })
        time.sleep(2)

    driver.quit()
    pd.DataFrame(ket_qua).to_excel("ketqua.xlsx", index=False)

if __name__ == "__main__":
    chay_tra_cuu()
