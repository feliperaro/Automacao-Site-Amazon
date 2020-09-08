from selenium import webdriver
import xlsxwriter as xlsxwriter

# Caminho do driver Chrome Selenium
driver = webdriver.Chrome('C:/chromedriver')


def abrePaginaNavegador(url):
    driver.get(url)


def pesquisaProduto(prod):
    search_box = driver.find_element_by_id("twotabsearchtextbox")
    search_box.send_keys(prod)
    search_box.submit()


def pegaNomes():
    return driver.find_elements_by_xpath("//span[@class='a-size-base-plus a-color-base a-text-normal']")


# Exibe nomes no terminal, caso necessário
def exibeNomes():
    nomes = pegaNomes()

    for nome in nomes:
        print(nome.text)


def pegaPrecos():
    return driver.find_elements_by_css_selector('span.a-price-whole')


# Exibe preços no terminal, caso necessário
def exibePrecos():
    precos = pegaNomes()
    for preco in precos:
        print("R${}".format(preco.text))


# Cria a planilha e/ou insere os dados
def montaExcel():
    nomes = pegaNomes()
    precos = pegaPrecos()

    workbook = xlsxwriter.Workbook('nomes_precos_iphones_amazon.xlsx')  # nome do arquivo
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 120)
    worksheet.set_column('B:B', 20)

    count = 0
    for preco in precos:
        # Insere todos os preços da página
        worksheet.write('B{}'.format(count + 1), 'R${}'.format(preco.text))
        count += 1
        # OBS: Algumas pesquisas podem não conter o preço, caso aconteca, os preços podem aparecer ao lado do produto errado

    count = 0
    for nome in nomes:
        # APENAS PRODUTOS COM "Iphone" NO NOME
        if "Iphone" in nome.text:
            worksheet.write('A{}'.format(count + 1), nome.text)
            count += 1

        # TODOS OS PRODUTOS DA PESQUISA
        #  worksheet.write('A{}'.format(count + 1), nome.text)
        #  count += 1

    workbook.close()


def fechaPagina():
    driver.quit()
