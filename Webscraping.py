from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Connect browser
browser = webdriver.Chrome()

#1. Get Dollar Rate
browser.get('https://www.google.com')
browser.find_element('xpath',
                     '/html/body/div[1]/div[3]/form/'
                     'div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Dollar rate')
browser.find_element('xpath',
                     '/html/body/div[1]/div[3]/form/div[1]/'
                     'div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
dollar_rate = browser.find_element('xpath',
                                   '//*[@id="knowledge-currency__updatable-data-column"]'
                                   '/div[1]/div[2]/span[1]').get_attribute('data-value')
#2. Get Euro Rate
browser.get('https://www.google.com')
browser.find_element('xpath',
                     '/html/body/div[1]/div[3]/form/'
                     'div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Euro rate')
browser.find_element('xpath',
                     '/html/body/div[1]/div[3]/form/div[1]/'
                     'div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
euro_rate = browser.find_element('xpath',
                                 '//*[@id="knowledge-currency__updatable-data-column"]'
                                 '/div[1]/div[2]/span[1]').get_attribute('data-value')
#3. Get Gold Rate
browser.get('https://www.melhorcambio.com/ouro-hoje')
gold_rate = browser.find_element('xpath',
                                   '//*[@id="comercial"]').get_attribute('value')
gold_rate = gold_rate.replace(',', '.')
#4. Import database

table = pd.read_excel('Produtos.xlsx')


#5. Recalculate the prices
# Change values of dollar, euro and gold
table.loc[table['Moeda'] == 'Dólar', 'Cotação'] = float(dollar_rate)

table.loc[table["Moeda"] == 'Euro', 'Cotação'] = float(euro_rate)

table.loc[table["Moeda"] == 'Ouro', 'Cotação'] = float(gold_rate)

# Update values of purchase

table['Preço de Compra'] = table['Preço Original'] * table['Cotação']

# Update Selling Price

table['Preço de Venda'] = table['Preço de Compra'] * table['Margem']

table['Preço de Venda'] = table['Preço de Venda'].map('R$ {:.2f}'.format)

# 6. Export the updated database

table.to_excel('Products-new.xlsx', index=False)

browser.quit()
