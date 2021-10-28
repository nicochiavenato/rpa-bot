from itdashboard.itdashboard import ItdashboardSelenium, ItdashboardExcel, ItdashboardPdf
from RPA.FileSystem import FileSystem
import os

# Limpa o diretório output
path_dir = os.path.dirname(os.path.abspath(__file__))
path_download = os.path.join(path_dir, './output')
botFilesystem = FileSystem()
botFilesystem.empty_directory(path_download)

# Entra no site e extrai os dados iniciais
botSelenium = ItdashboardSelenium()
botSelenium.first_page()
botSelenium.dive_in()
departaments_name, departaments_expense, selected_agency = botSelenium.get_all_expenses()

# Cria o Workbook e renomeia a planilha padrão
botExcel = ItdashboardExcel()
botExcel.create_workbook()
botExcel.rename_worksheet('Sheet', 'Agências')
# Preenche a planilha com os dados iniciais extraidos
botExcel.fill_sheet(departaments_name, departaments_expense)

# Entra na página de uma agência definida em constants.py
botSelenium.click_element_when_visible(selected_agency)
# Expande e captura a tabela
botSelenium.expand_table()
header, table = botSelenium.capture_table()

# Cria nova planilha
botExcel.create_worksheet('Individual Investments')
botExcel.set_active_worksheet('Individual Investments')
# Preenche a planilha com a tabela da agência específicada
botExcel.fill_sheet_with_table(header, table)
# Salva o Workbook
botExcel.save_workbook_with_path()

# Retorna links da tabela para download dos PDFs
links_elements = botSelenium.get_links_on_table()
if links_elements != '':
    # Download dos PDFs
    links_text = botSelenium.download_pdfs(links_elements)
    ################################ BONUS ################################
    # Lê a Seção A dos PDFs baixados
    names_investment = []
    uiis = []
    path_dir = os.path.dirname(os.path.abspath(__file__))
    for pdf_name in links_text:
        path_download = os.path.join(path_dir, f'./output/{pdf_name}.pdf')
        botPdf = ItdashboardPdf(path_download)
        # Adiciona os campos dos PDFs nos arrays names_investment e uiis
        botPdf.get_section_a(names_investment, uiis)
    # Faz a busca e comparação com a planilha
    botExcel.search_from_pdf(names_investment, uiis)
    ################################ /BONUS ###############################

# Fecha o navegador
botSelenium.close_all_browsers()
# Fecha o Workbook
botExcel.close_workbook()
