import logging
from datetime import timedelta, datetime as dt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import os
import time
import tkinter as tk
from tkinter import messagebox


def realizar_pesquisas_e_download(data_inicial, data_final, caminho_historico):
    logging.info("Iniciando a função realizar_pesquisas_e_download.")
    logging.info(f"Período da pesquisa: {data_inicial} a {data_final}")

    data_final_dt = dt.strptime(data_final, '%d/%m/%Y')
    data_inicial_dt = dt.strptime(data_inicial, '%d/%m/%Y')
    dif_dias = (data_final_dt - data_inicial_dt).days
    data_meio_dt = data_inicial_dt + timedelta(days=dif_dias // 2)
    data_meio = data_meio_dt.strftime('%d/%m/%Y')
    data_meio_1 = (data_meio_dt + timedelta(days=1)).strftime('%d/%m/%Y')    

    sieg_principal = "https://cofre.sieg.com/home.aspx"
    nfce = "https://cofre.sieg.com/pesquisa-avancada-nfce"
    nfe = "https://cofre.sieg.com/pesquisa-avancada"
    cfe = "https://cofre.sieg.com/pesquisa-avancada-cfe"
    cte = "https://cofre.sieg.com/pesquisa-avancada-cte"

    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": caminho_historico,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--start-maximized")
    try:
        driver = webdriver.Chrome(options=options)        
    except Exception as e:
        logging.exception("Erro ao iniciar o navegador. Verifique se o ChromeDriver está instalado corretamente e compatível com a versão do navegador.")
        
        root = tk.Tk()
        root.withdraw()  # Esconde a janela principal
        messagebox.showerror("Erro ao iniciar navegador", "Falha ao iniciar o Chrome. Verifique o ChromeDriver.")
    
        return


    try:  
        
        def preencher_pesquisa_avancada(valor, valor2):
            logging.info(f"Preenchendo pesquisa: {valor} até {valor2}")
            
            # Aguarda e seleciona "EmissionDate"
            combobox_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "fields"))
            )
            Select(combobox_element).select_by_value("EmissionDate")

            # Aguarda e seleciona "$gte"
            condicao_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "conditions"))
            )
            Select(condicao_element).select_by_value("$gte")

            # Aguarda campo de texto
            campo_digitavel = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "form-control.input-search.EmissionDate"))
            )
            campo_digitavel.clear()
            campo_digitavel.send_keys(valor)

            # Botão "addfields"
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "addfields.btn.btn-danger"))
            ).click()

            # Segunda condição "EmissionDate"
            segundo_campo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='optfield'][last()]/select[@class='fields form-control']"))
            )
            Select(segundo_campo).select_by_value("EmissionDate")

            # Segunda condição "$lte"
            segunda_condicao = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='optfield'][last()]/select[@class='conditions form-control']"))
            )
            Select(segunda_condicao).select_by_value("$lte")

            # Campo final de data
            campo_data_final = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "txtDate2"))
            )
            campo_data_final.clear()
            campo_data_final.send_keys(valor2)

            time.sleep(2)  # se puder, substitua por espera ativa (veja nota abaixo)

            # Botão "Pesquisar"
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "btn.btn-success"))
            ).click()

            # Botão "Marcar Todos"
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@class='dt-button btn btn-default margin-right-button esconder-mobile buttons-select-all']/span[text()='Marcar Todos']"))
            ).click()

            # Botão de exportar Excel
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "excel-export-btn"))
            ).click()

            logging.info("Pesquisa enviada com sucesso.")

        def aguardar_download(novo_nome=None):
            arquivos_anteriores = [arquivo for arquivo in os.listdir(caminho_historico) if arquivo.endswith(('.xls', '.xlsx'))]
            tempo_inicio = time.time()

            while time.time() - tempo_inicio < 900:
                arquivos_atuais = [arquivo for arquivo in os.listdir(caminho_historico) if arquivo.endswith(('.xls', '.xlsx'))]
                novos = list(set(arquivos_atuais) - set(arquivos_anteriores))
                if novos:
                    nome_original = novos[0]
                    logging.info(f"Download concluído: {nome_original}")
                    if novo_nome:
                        os.rename(os.path.join(caminho_historico, nome_original), os.path.join(caminho_historico, novo_nome))
                        logging.info(f"Arquivo renomeado para: {novo_nome}")
                    return
                time.sleep(2)
            logging.warning("Tempo de espera excedido. Download pode não ter sido concluído.")

        def realizar_login():
            logging.info("Acessando página de login.")
            driver.get(sieg_principal)

            try:
                # Aguarda os campos de login estarem presentes
                campo_email = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "txtEmail"))
                )
                campo_senha = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "txtPassword"))
                )
                botao_login = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "btnSubmit"))
                )

                campo_email.clear()
                campo_email.send_keys("fisc03@exatta.com.br")
                campo_senha.clear()
                campo_senha.send_keys("mrm310144")
                botao_login.click()

                logging.info("Login realizado com sucesso, aguardando confirmação...")

                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "h1"))
                )
                logging.info("Login confirmado, página carregada corretamente.")
                return True

            except TimeoutException:
                logging.error("Falha no login ou redirecionamento para página de erro.")
                return False

        
        if not realizar_login():
            logging.error("Login falhou. A automação será interrompida.")
            driver.quit()
            return

        def executar_pesquisa(nome, url, renomear=None):
            logging.info(f"Iniciando pesquisa: {nome}")
            # Esperar até a URL mudar
            url_atual = driver.current_url
            
            driver.get(url)
            
            try:
                WebDriverWait(driver, 10).until(lambda d: d.current_url != url_atual)
                logging.info("A URL mudou com sucesso.")
            except TimeoutException:
                logging.warning("A URL não mudou após o timeout.")

            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "fields"))
                )
                logging.info(f"Página de {nome} carregada com sucesso.")
            except Exception as e:
                logging.error(f"Erro ao aguardar carregamento da página de {nome}: {e}")
                try:
                    with open(f"erro_carregamento_{nome}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    logging.info(f"HTML da página de erro '{nome}' salvo como 'erro_carregamento_{nome}.html'")
                except Exception as ex:
                    logging.warning(f"Falha ao salvar HTML da página de erro {nome}: {ex}")
                raise

            if dif_dias <= 15:
                preencher_pesquisa_avancada(data_inicial, data_final)
                aguardar_download(renomear)
            else:
                preencher_pesquisa_avancada(data_inicial, data_meio)
                aguardar_download(renomear if not renomear else renomear.replace(".xlsx", "_parte1.xlsx"))

                driver.get(url)

                preencher_pesquisa_avancada(data_meio_1, data_final)
                aguardar_download(renomear if not renomear else renomear.replace(".xlsx", "_parte2.xlsx"))

            logging.info(f"Pesquisa {nome} finalizada.")

        # Chamando a função de pesquisa
        executar_pesquisa("NFC-e", nfce)
        executar_pesquisa("NF-e", nfe)
        executar_pesquisa("CF-e", cfe)
        executar_pesquisa("CT-e", cte, "CTE.xlsx")

    except Exception as e:
        logging.exception("Erro na função realizar_pesquisas_e_download:")
        try:
            with open("pagina_erro.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logging.info("Página de erro salva em 'pagina_erro.html'")
        except Exception as ex:
            logging.warning(f"Falha ao salvar HTML da página: {ex}")
        
    finally:
        driver.quit()
        logging.info("Navegador fechado e função concluída.")
