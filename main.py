from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv

from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(ChromeDriverManager().install())
url = "https://apps.mesmi.internal.ericsson.com//MIReports/SerialNumberLog?"

with open('Book1.csv', newline='') as csv_file_read:
    csv_reader = csv.reader(csv_file_read)
    csv_writer = csv.writer(open("RESULT.csv", 'w', newline=''))

    line_count = 0
    try:
        for row in csv_reader:

            driver.get(url)
            elem = driver.find_element_by_id("txtSerialNumber")
            elem.send_keys(Keys.BACKSPACE * 15)
            elem.send_keys(row[0])
            elem.send_keys(Keys.ENTER)

            TestPointCode = ""
            Value = ""
            error_codes = []

            product_number = ''

            try:

                elem = WebDriverWait(driver, 500).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "tr[role=row]"))
                )
                trow = driver.find_elements_by_css_selector("tr[role=row]")
                buttons = driver.find_elements_by_css_selector("span.misns-expand-icon")

                for i in range(len(trow)):
                    # print(len(trow[i].text.split('KRF')))

                    if len(trow[i].text.split('KRF ')) == 2:
                        product_number = 'KRF ' + trow[i].text.split('KRF ')[1]
                        # product_number = product_number.split(' ').join('')
                        # print(str(trow[i].text.split('KRF')))
                        break

                # print('\nPRODUCT NUMBER: KRF', product_number)

                for i in range(len(trow)):
                    if trow[i].text[:17] == "Functional Tested":
                        opened_test_info_div = buttons[i].find_element_by_xpath("..")
                        buttons[i].click()
                        span = WebDriverWait(opened_test_info_div, 3000).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.misns-log-info span"))
                        )

                error_code_spans = driver.find_elements_by_css_selector("div.misns-log-info span")
                for i in range(len(error_code_spans)):
                    error_message = ""
                    if error_code_spans[i].text[:13] == "TestPointCode":
                        TestPointCode = error_code_spans[i].text.split("TestPointCode: ")[1]
                    if error_code_spans[i].text[:5] == "Value":
                        Value = error_code_spans[i].text.split("Value: ")[1]
                        # print(float(Value) > -125)

                        if Value[0] == '-' and float(Value) > -125 and (TestPointCode[:6] == "IM3_C1" or TestPointCode[:6] == "IM3_C3"):
                            TestPointCode = 'RX' + 'ABCDABCDABCDABCD'[int(TestPointCode.split(".")[1]) - 1]
                            error_message = TestPointCode + " " + Value
                            error_codes.append(error_message)

                        # elif Value != "0" and Value != "101" and Value[0] == '-' and TestPointCode[:2] == "IM":
                        elif Value[0] == '-' and float(Value) > -125 and TestPointCode[:2] == "IM":
                            TestPointCode = 'RX' + TestPointCode.split(".")[0][2]
                            error_message = TestPointCode + " " + Value
                            error_codes.append(error_message)

                        elif Value[0] == '-' and float(Value) > -125 and TestPointCode[:2] == "12":
                            TestPointCode = "RX" + TestPointCode[2]
                            error_message = TestPointCode + " " + Value
                            error_codes.append(error_message)

                print(line_count + 1, '-', ''.join(product_number.split(' ')), row[0], " error_codes: ", error_codes)

            except ElementNotInteractableException:
                print("Error")

            if len(error_codes):
                csv_writer.writerow([''.join(product_number.split(' ')), row[0], error_codes[-1]])
            else:
                csv_writer.writerow([''.join(product_number.split(' ')), row[0], '-'])

            line_count += 1
        print(f'Processed {line_count} lines.')
    finally:
        driver.quit()
