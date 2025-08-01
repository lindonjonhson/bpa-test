#   Scraping related libs

import pandas as pd
import requests
import json
from dotenv import load_dotenv
import os

#   Process related libs

import subprocess
import time
import pyautogui
# import pytesseract


products_info = []

def read_xlsx(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f'Erro ao conectar via RDP: {e}')

def get_api_key():
    try:
        load_dotenv()
        return os.getenv("SERPAPI_API_KEY")
    except Exception as e:
        print(f'Erro ao importar API KEY: {e}')

def get_user_info():
    try:
        load_dotenv()
        user_info = {
            'user': os.getenv('USER'),
            'pw': os.getenv('PW')
        }
        return user_info
    except Exception as e:
        print(f'Erro ao importar informacoes do usuario: {e}')

def search(query):
    try:
        params ={
            'api_key': get_api_key(),
            'engine': 'amazon',
            'k': query,   #   Query
            'amazon_domain': "amazon.com.br",   #   Specifying the store's location
            # 's': 'price-asc-rank' # Looks up the lowest priced items, but it would not be a good idea as the lowest items might not be the desired ones
        }

        search = requests.get('https://serpapi.com/search', params=params)

        if not search.json().get('organic_results', [])[0].get('extracted_price'):
            raise ValueError('Value not available')

        product_info = {
            'product': query,
            'description': search.json().get('organic_results', [])[0].get('title'),
            'value': search.json().get('organic_results', [])[0].get('extracted_price'),
            'status': 'Uncompleted',
            'obs': 'Data scraped, but not registered'
        }

        return product_info
    except Exception as e:
        print(f'Erro ao procurar item: {e}')

        product_info = {
            'product': query,
            'description': search.json().get('organic_results', [])[0].get('title'),
            'value': 0,
            'status': 'Error',
            'obs': 'Error when scraping, possibly empty values'
        }
        # print(product_info)

def search_products(products):
    try:
        for product in products['Produto']:
            # print(product)
            prod_info = search(product)
            # print(prod_info)
            products_info.append(prod_info)
    except Exception as e:
        print(f'Erro ao procurar itens: {e}')

def connect():
    try:
        process = subprocess.Popen(["mstsc", "bpa.rdp"])

        #   Wait for the session to load 
        time.sleep(7)

        return process
    except Exception as e:
        print(f'Erro ao tentar connectar a máquina remota: {e}')

def login():
    try:
        #   Opening the application
        pyautogui.doubleClick(x=33, y=409)

        #   Wait for the application to start -- Change the value in case the application is taking longer to start
        time.sleep(7)  

        user_info = get_user_info()

        #   Loggin in
        print('Logging in... \n')
        pyautogui.typewrite(user_info['user'])
        pyautogui.press('tab')
        pyautogui.typewrite(user_info['pw'])
        pyautogui.press('enter')
        print('Logged In!\n')

        #   STRPDM

        pyautogui.typewrite('strpdm')
        pyautogui.press('enter')

        #   3

        pyautogui.typewrite('3')
        pyautogui.press('enter')

        #   QCPPSRC

        pyautogui.typewrite('qcppsrc')
        pyautogui.press('enter')
    except Exception as e:
        print(f'Erro ao tentar abrir aplicacao: {e}')

def close_app():
    try:
        pyautogui.hotkey('alt', 'f4')

        #   Press enter to confirm
        pyautogui.press('enter')
    except Exception as e:
        print(f'Erro ao tentar fechar aplicacao: {e}')

def add_item(product):
    try:
        print('Inserindo produto: ')
        print(product['product'])
        #   Validating if the product's status
        if product['status'] in ('error','failed'):
            return

        #   Creating a new item


        pyautogui.press('f6')

        #   Title
        pyautogui.typewrite(product['product'].replace(" ", ""))
        pyautogui.press('tab')

        #   Deleting the placeholder
        pyautogui.press('right', presses=5)
        pyautogui.press('backspace', presses=5)
        #   TXT
        pyautogui.typewrite('TXT')

        pyautogui.press('tab')

        #   Value
        pyautogui.typewrite(f"Valor {product['value']}")

        #   Next page
        pyautogui.press('enter')

        #   Item description
        pyautogui.typewrite(product['description'][:70])

        #   Confirm
        pyautogui.press('enter')

        pyautogui.press('f3')

        #   Save changes and go back to the main page

        pyautogui.press('enter')
        product['status'] = 'Success'
        product['obs'] = 'Item registered successfully'

        name = product['product']
        print(f'Produto {name} cadastrado!\n')

    except Exception as e:
        print(f'Erro ao tentar cadastrar item: {e}')

#   Reading file
df = read_xlsx('research.xlsx')
print(df)

#   Searching the products online
search_products(df)

#   Connecting to the remote machine
process = connect()

#   Logging in
login()

#   Adding items
print('Adding items...')
for product in products_info:
    add_item(product)

print('Closing app...')
close_app()

print('Closing connection to RM...')
process.terminate()

#   Updating status
for index, product in enumerate(products_info):
    df['Valor'][index] = product['value']
    df['Descrição'][index] = product['description']
    df['Status_Processamento'][index] = product['status']
    df['Obs'][index] = product['obs']

#   Saving to excel spreadsheet
df.to_excel("research.xlsx", index=False)