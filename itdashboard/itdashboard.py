import itdashboard.constants as const
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from selenium.webdriver.firefox.options import Options
from PyPDF2 import PdfFileReader
import time
import os
import re


class ItdashboardSelenium(Selenium):

    def first_page(self, wait=20):
        # Definições iniciais do navegador
        path_dir = os.path.dirname(os.path.abspath(__file__))
        path_download = os.path.join(path_dir, '../output/')
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.dir", path_download)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
        options.set_preference("pdfjs.disabled", True)
        self.open_browser(const.BASE_URL, options=options)
        self.set_browser_implicit_wait(wait)

    def dive_in(self):
        # Clica no elemento inicial requisitado
        self.click_element_when_visible('css:a[href="#home-dive-in"]')

    def get_all_expenses(self):
        # Busca o nome e gastos dos departamentos
        departaments_name_seal = self.find_elements('class=seals')
        departaments_name = []
        selected_agency = ''
        for dep in departaments_name_seal:
            departaments_name.append(dep.get_attribute('alt')[11:])
            # Encontra a agência escolhida
            if const.SELECT_AGENCY in dep.get_attribute('alt')[11:]:
                selected_agency = dep

        departaments_expense_full = self.find_elements('class=h1.w900')
        departaments_expense = []
        for i in range(len(departaments_expense_full)):
            departaments_expense.append(departaments_expense_full[i].text)

        return [departaments_name, departaments_expense, selected_agency]

    def expand_table(self):
        # Expande a tabela para mostrar todos os registros
        select = self.find_element('name=investments-table-object_length')
        self.select_from_list_by_value(select, '-1')
        # Tempo para recarregar tabela com todos os dados
        time.sleep(10)

    def capture_table(self):
        table_aux = []
        table = []
        header = []
        # Captura o cabeçalho da tabela
        headers = self.find_elements('css:th')
        for el in headers:
            if el.text != '':
                header.append(el.text)
        # Encontra o número de linhas da tabela
        number_entries = self.find_element('id=investments-table-object_info').text
        number_entries_split = number_entries.split()
        number_entries_total = int(number_entries_split[5])
        # Captura os dados da tabela
        count = 0
        for i in range(3, number_entries_total+3):
            for j in range(1, 8):
                table_aux.append(self.get_table_cell('id=investments-table-object', i, j))
                count += 1
                if count == 7:
                    table.append(table_aux)
                    table_aux = []
                    count = 0
        return [header, table]

    def get_links_on_table(self):
        # Captura possíveis links para download dos PDFs
        try:
            links_elements = self.find_elements('css:td a')
        except:
            links_elements = ''
        return links_elements

    def download_pdfs(self, links_elements):
        actual_url = self.get_location()
        # Cria lista de links a partir dos elementos
        links = []
        links_text = []
        for element in links_elements:
            links.append(actual_url+'/'+element.text)
            # Para ser usado na leitura do PDF
            links_text.append(element.text)
        for link in links:
            self.go_to(link)
            time.sleep(5)
            pdf_link = self.find_element('css:div[id="business-case-pdf"] a')
            self.click_element_when_visible(pdf_link)
            time.sleep(15)
        return links_text


class ItdashboardExcel(Files):
    def fill_sheet(self, departaments_name, departaments_expense):
        # Preenche a planilha Agências com os valores encontrados
        for i in range(len(departaments_name)):
            self.set_cell_value(i+1, 'A', departaments_name[i])
            self.set_cell_value(i+1, 'B', departaments_expense[i])

    def fill_sheet_with_table(self, header, table):
        # Preenche o cabeçalho da planilha Individual Investments
        for i in range(len(header)):
            self.set_cell_value(1, i+1, header[i])
        # Preenche os dados da planilha Individual Investments
        for i in range(len(table)):
            for j in range(len(table[0])):
                self.set_cell_value(i+2, j+1, table[i][j])

    def save_workbook_with_path(self):
        path_dir = os.path.dirname(os.path.abspath(__file__))
        path_workbook = os.path.join(path_dir, '../output/itdashboard.xlsx')
        self.save_workbook(path_workbook)

    def search_from_pdf(self, names_investment, uiis):
        investment_title = 'init'
        row = 2
        dic_table = {}
        # Busca na planilha até encontrar um campo vazio
        while investment_title is not None:
            investment_title = self.get_cell_value(row, 3)
            if investment_title is not None:
                dic_table[investment_title.lower()] = self.get_cell_value(row, 1)
            row += 1
        # Compara os dados do PDF com os dados da planilha
        for i in range(len(names_investment)):
            if names_investment[i].lower() in dic_table:
                if uiis[i] == dic_table[names_investment[i].lower()]:
                    print('############# [campos iguais] #############')
                    print('-> Campos PDF: ', names_investment[i].lower(), ' - ', uiis[i])
                    print('-> Campos XLS: ', names_investment[i].lower(),
                          ' - ', dic_table[names_investment[i].lower()])


class ItdashboardPdf(PdfFileReader):
    def get_section_a(self, names_investment, uiis):
        text_page_six = self.getPage(5).extractText()
        # Realiza a busca utilizando Expressão Regular
        name_investment = re.findall(r"Name of this Investment:[\s\n]*(.*)\s", text_page_six)
        uii = re.findall(r"UII (.*)", text_page_six)
        names_investment.append(name_investment[0][:-2])
        uiis.append(uii[0])
