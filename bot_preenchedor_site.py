#capta info de uma planilha e preenche em um site estilo formulário
from selenium import webdriver #pegar o navegador
from selenium.webdriver.common.by import By #pegar o By para selecionar campos em um site
from time import sleep 
import openpyxl

#abrir o chrome
driver = webdriver.Chrome()
driver.get('https://contabilidade-devaprender.netlify.app/') #entrar no link de login do site
sleep(2)
# para solicitar uma tag especifica usar XPATH(tag[@atributo='valor']) 
#ex: (By.XPATH,"//input[@id='Email']")
#quando tiver id, usar apenas (by.ID,'nome do id')

#preencher o campo EMAIL do login com o email do usuário
email = driver.find_element(By.ID,'email')
sleep(1)
email.send_keys('admin@contabilidade.com')

#preencher o campo SENHA do login com a Senha do usuário
senha = driver.find_element(By.ID,'senha')
sleep(1)

senha.send_keys('contabilidade123456')
entrar= driver.find_element(By.XPATH,"//button[@id='Entrar']")
sleep(2)
entrar.click()
sleep(5)

#entrar na planilha para ler as informaçoes

planilha =openpyxl.load_workbook('bot_preenchedor-site/empresas.xlsx')
pagina_certa = planilha['dados empresas'] #pagina específica da planilha

for linha in pagina_certa.iter_rows(min_row=2, values_only= True): #iterar sobre cada coluna de cada linha (começando da linha 2)
    nome_empresa, Email, telefone, endereco, cnpj, area, qtd_funcionarios, data_fundacao = linha
    #para cada linha, preencher as informaçoes com os valores aquiridos
    driver.find_element(By.ID,'nomeEmpresa').send_keys(nome_empresa)
    sleep(1)
    driver.find_element(By.ID,'emailEmpresa').send_keys(Email)
    sleep(1)
    driver.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(1)
    driver.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(1)
    driver.find_element(By.ID,'cnpj').send_keys(cnpj)
    sleep(1)
    driver.find_element(By.ID,'areaAtuacao').send_keys(area)
    sleep(1)
    driver.find_element(By.ID,'numeroFuncionarios').send_keys(qtd_funcionarios)
    sleep(1)
    driver.find_element(By.ID,'dataFundacao').send_keys(data_fundacao)
    sleep(1)
    driver.find_element(By.ID,'Cadastrar').click()
    sleep(5)
    
    


