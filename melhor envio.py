from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
import openpyxl
from selenium.common.exceptions import NoSuchElementException   


driver = webdriver.Chrome(executable_path=r"C:\Users\Elisa\Documents\Python\chromedriver.exe")
book = openpyxl.load_workbook('Teste 1.xlsx')
sheet = book.active

driver.get("https://melhorenvio.com.br/calculadora")
driver.find_element_by_name("username").send_keys("41414723873")
driver.find_element_by_name("password").send_keys("asdf1234")
driver.find_element_by_name("password").send_keys(Keys.RETURN)

i=2
a=66

while i < 100:

        cep = sheet.cell(i,1).value
        nome = sheet.cell(i,2).value
        numero = sheet.cell(i,3).value
        comp = sheet.cell(i,4).value

        time.sleep(10)

        driver.find_element_by_name("iptCepDestino").send_keys(cep)
        #driver.find_element_by_name("iptAltura-21").send_keys("11")
        #driver.find_element_by_name("iptLargura-21").send_keys("5")
        #driver.find_element_by_name("iptComprimento-21").send_keys("20")
        #driver.find_element_by_name("iptPeso-21").send_keys("0,8")
        
        #driver.find_element_by_name("iptValorSegurado").send_keys(1000)
        driver.find_element_by_name("iptCepDestino").send_keys(Keys.RETURN)
    
        time.sleep(3)

        #if driver.find_element(By.CLASS_NAME, 'btnQuero').size != 0
       #     driver.find_element_by_name("iptValorSeguradoModal").send_keys(5000)

       # else

        btn_more = driver.find_element(By.CLASS_NAME, 'btnQuero')
        btn_more.click()

        #btn_more = driver.find_element(By.XPATH, "//*[ contains (text(), 'Selecionar' )]")
        # btn_more = driver.find_element(By.XPATH, //*[@id="secCalculator"]/div/form/div[2]/div[2]/div/div[3]/ul[2]/li[1]/ul/li[6]/span/button)

        time.sleep(2)

        
        s = "iptNome-"
        d = "iptEnderecoNum-"
        g = "iptComplemento-"
        
    

        driver.find_element_by_name(s+str(a)).send_keys(nome)
        driver.find_element_by_name(d+str(a)).send_keys(numero)
        driver.find_element_by_name(g+str(a)).send_keys(comp)

        a = a + 55

        btn_more1 = driver.find_element(By.CLASS_NAME, 'btnAddCart')
        btn_more1.click()

        time.sleep(2)
        btn_more2 = driver.find_element(By.CLASS_NAME, 'btnMore')
        btn_more2.click()
        
       

        i = i+ 1
