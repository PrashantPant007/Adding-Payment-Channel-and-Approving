"""Program to automate data"""



"""Step 1: Importing necessary libraries webdriver for browser, openpyxl to read,write process excel data"""
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import chromedriver_autoinstaller
chromedriver_autoinstaller.install()
import openpyxl


"""Making a variable for a file directory"""
filepath = ("D:/FirstPay Drive/Automation/merchants.xlsx")

"""Openpyxl library to load filepath"""
excelfile = openpyxl.load_workbook(filepath) 

"""Reading active sheet from excel"""
sheet_obj = excelfile.active

"""Maximum rows and columns in sheet"""
rows = sheet_obj.max_row
cols = sheet_obj.max_column


"""Browser loading with webdriver"""
driver = webdriver.Chrome()
driver.get('https://spay.firstpay.com.np/login')
#parent_handle = driver.current_wndow_handle 
driver.implicitly_wait(8)                                                       #Wait for loading
    

"""Providing username and password"""    
driver.find_element_by_xpath("//input[@id='username']").send_keys("suman.bartaula")
driver.find_element_by_xpath("//input[@id='password']").send_keys("Welcome2")
    
 
"""Waiting to read captcha and type captcha"""
try:
    element = WebDriverWait(driver,10).until(EC.presence_of_element_located((BY.LINK_TEXT, "Sign in")))
    element.click()
except:
    time.sleep(10)
    driver.find_element_by_xpath("//button[contains(text(),'Sign in')]").click() #Click sign in button
     
        
time.sleep(3)     
driver.find_element_by_xpath("//span[contains(text(),'Documents')]").click()  
driver.maximize_window()
#parent_handle = driver.current_wndow_handle 
time.sleep(2)  
#driver.get("https://spay.firstpay.com.np/document/merchant")    

driver.find_element_by_xpath("//body[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/ul[1]/li[2]/ul[1]/li[2]/span[1]/span[1]").click()  #Click on Merchant


"""mrow for iterating rows"""
mrow = 1

while (mrow<rows+1):
    
    
     
     #time.sleep(3)
  
     #OM = driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/span[1]/div[6]/div[2]/div[1]/div[1]/input[1]") #Click on add an ordinary merchant
      
         
     for mrow in range(2,rows+1): 
        
                                                 #loop for rows
     
         #OM.click()#driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/span[1]/div[6]/div[2]/div[1]/div[1]/input[1]").send_keys(Merchant_Name)
                                                              #Clicking add an ordinary merchant depending upon the rows
         for mcols in range(1,cols+1):
             
                                      #Data inside a single rows depending upon columns
         
          
             Parent_Agent = sheet_obj.cell(mrow,1).value
             Merchant_Country = sheet_obj.cell(mrow,2).value
             Merchant_City = sheet_obj.cell(mrow,3).value
             Merchant_Name = sheet_obj.cell(mrow,4).value
             Abr_of_Merchant_Name = sheet_obj.cell(mrow,5).value
             Service_Hotline = sheet_obj.cell(mrow,6).value
             Merchant_Currency = sheet_obj.cell(mrow,7).value
             Settlement_Rules = sheet_obj.cell(mrow,8).value
             Major_product_Services = sheet_obj.cell(mrow,9).value
             Contact_Name = sheet_obj.cell(mrow,10).value
             Contact_Phone = sheet_obj.cell(mrow,11).value
             Email_Address = sheet_obj.cell(mrow,12).value
             Address = sheet_obj.cell(mrow,13).value
             Payment_type = sheet_obj.cell(mrow,15).value
             Merchant_Discount_Rate_WeChatPay = sheet_obj.cell(mrow,16).value
             Merchant_Category_Code_WeChatPay = sheet_obj.cell(mrow,17).value
             Business_Category = sheet_obj.cell(mrow,18).value
             Merchant_Type = sheet_obj.cell(mrow,19).value 
             Registered_Certificate_ID = sheet_obj.cell(mrow,20).value
             Certificate_Validity = sheet_obj.cell(mrow,21).value
             Certificate_Photo = sheet_obj.cell(mrow,22).value
             Business_Type = sheet_obj.cell(mrow,23).value
             If_Offline_Store_Address = sheet_obj.cell(mrow,25).value
             Store_Photos = sheet_obj.cell(mrow,26).value
             Merchant_Discount_Rate = sheet_obj.cell(mrow,27).value
             Merchant_ID_for_PaymentUPI = sheet_obj.cell(mrow,28).value
             Merchant_Category_CodeUPI = sheet_obj.cell(mrow,29).value
             Terminal_IDUPI = sheet_obj.cell(mrow,30).value
             Bank_Account_Number = sheet_obj.cell(mrow,31).value
             Account_Name = sheet_obj.cell(mrow,32).value
             Bank_Name = sheet_obj.cell(mrow,33).value
             Branch_Name = sheet_obj.cell(mrow,34).value 
             Account_Type = sheet_obj.cell(mrow,35).value
             Swift_Code = sheet_obj.cell(mrow,36).value
               
                  
             """Page 1: Merchant basic information""" 
             try:
                 time.sleep(2)
                 driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/span[1]/div[6]/div[2]/div[1]/div[1]/input[1]").send_keys(Merchant_Name) 
                 time.sleep(2)  
                 #driver.find_element_by_xpath("//input[@id='']").send_keys(Merchant_City)               
                  #time.sleep(2)  
                 driver.find_element_by_xpath("//button[contains(text(),'Search')]").click()
                 time.sleep(2)
                  #clicking on adding payment channel
                 driver.find_element_by_xpath("//tbody/tr[1]/td[11]/div[1]/i[4]").click()
                  
                 #time.sleep(2)
                 #driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default div.ant-modal-content.react-draggable:nth-child(2) strong.cursor > div.ant-modal-header").click()
                 #time.sleep(2)
                 
                 #driver.find_element_by_xpath("//button[contains(text(),'OK')]").click()
                 #driver.find_element_by_xpath("//div[contains(text(),'Payment Channel1: FirstPay-Union Pay-9')]")
                 #time.sleep(2)
                 #driver.find_element_by_xpath("/html[1]/body[1]/div[6]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/i[1]").click()
                 #driver.find_element_by_id("payCenterInfo[0].merchantNameEn").click() 
                 time.sleep(2)                     
                    #click on rate type(UPI)
                 driver.find_element_by_xpath("//div[contains(text(),'Please select')]").click()
                 time.sleep(2)
                 #driver.find_element_by_xpath("/html[1]/body[1]/div[6]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]").click()
                 #time.sleep(2)
                 #click on flat rate (QR)
                 driver.find_element_by_xpath("//li[contains(text(),'Flat rate(QR)')]").click()
                 time.sleep(2)
                 #click on MDR
                 driver.find_element_by_id("payCenterInfo[0].payRateFt").send_keys(Merchant_Discount_Rate)
                 time.sleep(2)
                 driver.find_element_by_id("payCenterInfo[0].merchantNameEn").send_keys(Merchant_Name)
                    #click on merchant name
                 #driver.find_element_by_xpath("/html[1]/body[1]/div[10]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/input[1]").click()    #send_keys(Merchant_Name)
                 time.sleep(2)
                    #merchant id for payment
                 driver.find_element_by_id("payCenterInfo[0].payMerchant").send_keys(Merchant_ID_for_PaymentUPI)
                 time.sleep(2)
                 driver.find_element_by_id("payCenterInfo[0].pid").send_keys(Terminal_IDUPI)
                 time.sleep(2)
                   #Merchant Category code
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-body div.popPaddingLeft div.addMch_module_paycenter:nth-child(1) div.addMch_module_paycenter_content div.swift_labelcomp.swift_formplus_label:nth-child(6) div.swift_labelcomp_cont div.formitem_wrap div.swiftBI_searchbar_oversea.bar_plus div.search_bar_outer_input_wrap div.swift_input_wrap div:nth-child(1) > input.swiftBI_input.primary.medium").click()
                 time.sleep(1)
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default.searchbar_oversea_pop div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-body div.searchbar_oversea_filter div.big_btm div.swift_labelcomp.primary.searchbar_single_input div.swift_labelcomp_cont div.swift_input_wrap div:nth-child(1) > input.swiftBI_input.reset.primary.medium").send_keys(Merchant_Category_CodeUPI)
                 time.sleep(1)
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default.searchbar_oversea_pop div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-body div.searchbar_oversea_filter div.big_btm > button.swiftBI_button.search.size_medium").click()
                 time.sleep(1)
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default.searchbar_oversea_pop div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-body div.searchbar_oversea_cont div.ant-table-wrapper.swiftBI_table.all_align_left div.ant-spin-nested-loading div.ant-spin-container div.ant-table.ant-table-default.ant-table-fixed-header.ant-table-scroll-position-left div.ant-table-content div.ant-table-scroll div.ant-table-body table:nth-child(1) tbody.ant-table-tbody:nth-child(2) tr.ant-table-row.ant-table-row-level-0 > td:nth-child(2)").click()
                 time.sleep(2)
                 
                 #click on wechat pay off button
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-body div.popPaddingLeft div.addMch_module_paycenter:nth-child(2) div.addMch_module_paycenter_content div.swift_labelcomp.swift_formplus_label.addMch_module_item:nth-child(8) div.swift_labelcomp_cont div.formitem_wrap div.swiftBI_radio_plus_wrap div.swiftBI_radio_plus.primary.size_small:nth-child(2) > i:nth-child(1)").click()
                 time.sleep(2)
                   
                    #click on okay button
                 driver.find_element_by_xpath("//button[contains(text(),'OK')]").click()
                 time.sleep(2)
                 #click on pending button
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.swiftBI_flex.vertical.full_height.shouldMinWidth div.swiftBI_flex_item.init_scroll div.hs_content div.ant-table-wrapper.swiftBI_table div.ant-spin-nested-loading div.ant-spin-container div.ant-table.ant-table-default.ant-table-scroll-position-left div.ant-table-content div.ant-table-body tbody.ant-table-tbody:nth-child(3) tr.ant-table-row.ant-table-row-level-0:nth-child(1) td:nth-child(10) p:nth-child(1) a.swiftBI_button.a_code > span:nth-child(1)").click()
                 time.sleep(2)
                 #click on approve button
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.ant-modal-wrap.swift_popup_center div.ant-modal.swift_popup.cont_bg_default div.ant-modal-content.react-draggable:nth-child(2) div.ant-modal-footer div:nth-child(1) > button.swiftBI_button.primary.size_medium:nth-child(2)").click()
                 time.sleep(2)
                 #crossing merchant name
                 driver.find_element_by_css_selector("body.gray_background:nth-child(2) div.swiftBI_flex.vertical.full_height.shouldMinWidth div.swiftBI_flex_item.init_scroll div.hs_content div.filter_wrap div.filter_bd div.filter_row.search_eara:nth-child(1) div.search_box div.swift_labelcomp.primary.common_searchBar:nth-child(6) div.swift_labelcomp_cont div.swift_input_wrap div:nth-child(1) span.input_extra > i.swiftBI_icon.icon-fail.iconfont.input_reset").click()
                 time.sleep(1)
                    
                    
                    
                 
                 
                 sheet_obj.cell(mrow,cols).value = "Completed" #When a single merchant data is successfully enrolled from excelsheet, it writes completed in excelsheet. 
                 excelfile.save(filepath)
                 time.sleep(3)
                  
                  #driver.close()
                  #driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[2]/button[1]").click()
                 #time.sleep(3)
             except Exception:
                 sheet_obj.cell(mrow,cols).value = "Exception Error"
                 excelfile.save(filepath)
                 time.sleep(2)
                 #driver.close()
                 continue
             except ValueError:
                 sheet_obj.cell(mrow,cols).value = "ValueError"
                 excelfile.save(filepath)
                 time.sleep(2)
                 #driver.close()
                 continue
                 #driver.find_element_by_xpath("/html[1]/body[1]/div[12]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/button[1]").click()
                  
                  
                  
                   
             mrow+=1 
                 #OM.click()   
                    
                 
                    #driver.quit()  
