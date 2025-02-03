from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


class CNPJScraper:
    """Classe para automatizar a consulta de CNPJs e salvar os dados em um arquivo Excel."""
    
    def __init__(self, site_url):
        options = webdriver.ChromeOptions()
        #options.add_argument("--headless")  # Executa sem abrir o navegador
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        self.site_url = site_url
    
    def acessar_site(self):
        self.driver.get(self.site_url)
        time.sleep(3)

    def buscar_dados_cnpj(self, cnpj):
        try:
            if self.driver.find_elements(By.LINK_TEXT, "Aceito os Termos"):
                input_box = self.driver.find_element(By.LINK_TEXT, "Aceito os Termos")
                input_box.click()
            input_box = self.driver.find_element(By.TAG_NAME, "input")
            input_box.clear()
            input_box.send_keys(cnpj)
         
           # input_box.send_keys(Keys.RETURN)
            time.sleep(5)  

            endereco = self.driver.find_element(By.XPATH, '//div[@class ="p4 bg--secondary print-border print-mr-2 print-grow-3"]/p').text
            razao_social = self.driver.find_element(By.XPATH, '//div[@class ="p4 print-border"]/p').text

            return [cnpj, razao_social, endereco]
        except Exception as e:
            print(f"Erro ao buscar o CNPJ {cnpj}: {e}")
            return [cnpj, "Erro", "Erro"]

    def fechar_navegador(self):
        self.driver.quit()


class ExcelHandler:
    """Classe para manipulação do arquivo Excel."""
    
    def __init__(self, filename="C:\Temp\dados_cnpj.xlsx"):
        self.filename = filename
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.sheet.append(["CNPJ", "Razão Social", "Endereço"])

    def adicionar_dados(self, dados):
        self.sheet.append(dados)

    def salvar_arquivo(self):
        self.workbook.save(self.filename)
        print(f"Dados salvos em '{self.filename}'.")


def main():
    site_url = "https://www.consultacnpj.com"
    cnpjs = ["00.000.000/0001-91", "90.400.888/0001-42"]

    scraper = CNPJScraper(site_url)
    scraper.acessar_site()

    excel = ExcelHandler()

    for cnpj in cnpjs:
        dados = scraper.buscar_dados_cnpj(cnpj)
        excel.adicionar_dados(dados)

    excel.salvar_arquivo()
    scraper.fechar_navegador()
    print("Automação concluída! 🚀")

if __name__ == "__main__":
    main()
