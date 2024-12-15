from selenium import webdriver
from datetime import datetime
import pandas as pd
from time import sleep
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoAlertPresentException

# Configurações globais
CHROMEDRIVER_PATH = "chromedriver-win64\chromedriver.exe"
SITE_URL = "https://cenegedrj.gpm.srv.br/index.php"
LOGIN_USUARIO = "fabricio.gama"
SENHA_USUARIO = "11543339735*"

class Apontamento:
    def __init__(self):
        """Inicializa o driver e configurações do navegador."""
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")  # Abre o navegador maximizado
        chrome_options.add_experimental_option("detach", True)  # Mantém o navegador aberto

        self.service = Service(CHROMEDRIVER_PATH)
        self.driver = webdriver.Chrome(service=self.service)
        self.wait = WebDriverWait(self.driver, 50)
        self.log_file = "robo_apontamento_log.txt"
        self._init_log()

    def _init_log(self):
        """Inicializa o arquivo de log."""
        with open(self.log_file, "w") as log:
            log.write(f"Log iniciado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log.write("=" * 50 + "\n")

    def log(self, message):
        """Registra uma mensagem no arquivo de log."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"[{timestamp}] {message}"
        
        # Evitar mensagens duplicadas
        if not hasattr(self, '_last_log_message') or self._last_log_message != log_message:
            with open(self.log_file, "a") as log:
                log.write(log_message + "\n")
            print(log_message)  # Também imprime no console
            self._last_log_message = log_message

    def login(self):
        """Realiza o login no site."""
        try:
            self.driver.get(SITE_URL)
            self.wait.until(EC.presence_of_element_located((By.ID, "idLogin"))).send_keys(LOGIN_USUARIO)
            self.wait.until(EC.presence_of_element_located((By.ID, "idSenha"))).send_keys(SENHA_USUARIO)
            self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "blogin"))).click()
            self.log("Login realizado com sucesso.")
        except TimeoutException:
            self.log("Erro ao realizar o login. Verifique as credenciais ou a conexão.")
            self.driver.quit()
            exit()

    def botao_serviço(self):

        """ Clicar no botão de serviço """
        try:
            self.driver.implicitly_wait(5)
            # Clicar no botão de serviço
            self.wait.until(EC.element_to_be_clickable((By.ID, "2000"))).click()
            self.log("Botão de serviço clicado.")
        except TimeoutException:
            self.log("Botão de serviço não encontrado.")
            # self.driver.quit()
            # exit()

    def _acessar_iframes_lateral(self):

        """Troca para os iframes laterais."""
        try:
            # Acessar iframe lateral
            iframe_lateral = self.wait.until(EC.presence_of_element_located((By.ID, "frame_lateral")))
            self.driver.implicitly_wait(5)
            self.driver.switch_to.frame(iframe_lateral)
            self.log("Mudança para o iframe lateral realizada com sucesso.")
            # Clicar nos itens do menu
            self.driver.find_element(By.ID, "jt80").click()
            self.driver.find_element(By.ID, "jt111").click()
            self.driver.find_element(By.ID, "jt112").click()
            self.log("Itens do menu clicados com sucesso.")
            self.driver.switch_to.default_content()  # Voltar ao contexto principal
        except TimeoutException:
            self.log("Erro ao acessar os iframes. Verifique a estrutura da página.")
            # self.driver.quit()
            # exit()

    def _acessar_iframes_central(self):

        """Troca para os iframes centrais."""
        try:
            # Alternar para o contexto padrão antes de acessar o iframe central
            self.driver.switch_to.default_content()
            iframe_central = self.wait.until(EC.presence_of_element_located((By.ID, "frame_central")))
            self.driver.switch_to.frame(iframe_central)
            self.log("Mudança para o iframe central realizada com sucesso.")
        except Exception as e:
            self.log(f"Mudança para o iframe central erro: {e}.")
            # self.driver.quit()
            # exit()

    def _acessar_iframes_secundarios(self):
        """Acessa os iframes secundários."""
        try:
            self.log("Tentando acessar o iframe secundário...")
            # Verifica se o iframe já está ativo
            if self.driver.find_element(By.ID, "frm_down").is_displayed():
                self.log("Iframe secundário já está ativo.")

                # Acessar iframe secundário
                iframe_servico = self.wait.until(EC.presence_of_element_located((By.ID, "frm_down")))
                self.driver.switch_to.frame(iframe_servico)
                self.log("Mudança para o iframe lateral realizada com sucesso.")
                return  
         
        except TimeoutException:
            self.log("Erro ao acessar os iframes. Verifique a estrutura da página.")
            self.driver.quit()
            exit()

    def preencher_cabecalho(self, row):
        try:
                  
            self._preencher_com_sugestao("inputString", row["inputString"], "autoSuggestionsList")
            # Pausa breve
            sleep(0.1)
            self._interagir_dropdown("contrato_chosen", row["contrato_chosen"])
            # Pausa breve
            sleep(0.1)
            self._interagir_dropdown("equipe_chosen", row["equipe_chosen"])
            # Pausa breve
            sleep(0.1)
            self._interagir_dropdown("tip_srv_chosen", row["tip_srv_chosen"])
            # Pausa breve
            sleep(0.1)
            self._interagir_dropdown("obras_chosen", row["obras_chosen"])
            # Pausa breve
            sleep(0.1)
            self._interagir_dropdown("cod_irr_chosen", str(row["cod_irr_chosen"]))
            # Pausa breve
            sleep(0.1)

            # Preenchimento de data
            self._preencher_campo_data_hora("dat_srv", row["dat_srv"].strftime("%d/%m/%Y"))
            # Pausa breve
            sleep(0.1)
            self._preencher_campo_data_hora("hr_inic", row["hr_inic"].strftime("%H:%M"))
            # Pausa breve
            sleep(0.1)
            self._preencher_campo_data_hora("dat_srv2", row["dat_srv2"].strftime("%d/%m/%Y"))
            # Pausa breve
            sleep(0.1)
            self._preencher_campo_data_hora("hr_fim", row["hr_fim"].strftime("%H:%M"))
            self.log("Cabeçalho preenchido com sucesso.")
        except Exception as e:
            self.log(f"Erro ao preencher o cabeçalho: {e}")
            self.driver.quit()
            exit()
        finally:
            # salvar
            apontamento.salvar()
            # Acessar iframe secundário
            apontamento._acessar_iframes_secundarios()

    def preencher_servico(self,row):
        """Preenche os dados do servico."""
        try:

            # Preencher o campo de texto
            self._interagir_dropdown("serv_chosen", str(row["serv_chosen"]))

            # Garantir que 'qtd' seja um número válido e formatá-lo com dois dígitos decimais
            if isinstance(row["qtd"], (int, float)):
                valor_qtd = f"{float(row['qtd']):.2f}"  # Formata o número para sempre ter 2 casas decimais
            else:
                raise ValueError(f"Valor inválido para 'qtd': {row['qtd']}")
            
            # Preencher o campo com o valor formatado
            self._preencher_campo_data_hora("qtd", valor_qtd)

            # inclui servico
            apontamento.incluir()
            
        except Exception as e:
            self.log(f"Erro ao preencher os dados: {e}")

    def _preencher_com_sugestao(self, campo_id, texto, suggestion_list_id):
        """Preenche um campo de texto e clica na sugestão correspondente."""
        try:

            # Localizar o campo de texto e inserir o texto
            campo_texto = self.wait.until(EC.element_to_be_clickable((By.ID, campo_id)))
            campo_texto.clear()
            campo_texto.send_keys(texto)
            self.log(f"Texto '{texto}' inserido no campo '{campo_id}'.")

            # Esperar pela lista de sugestões aparecer
            sugestao = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f"#{suggestion_list_id} li"))
            )
            self.log(f"Texto '{texto}' na lista '{suggestion_list_id}'.")

            # Pausa breve
            sleep(0.5)
            try:
                # Clicar na primeira sugestão
                sugestao.click()
                self.log(f"Primeira sugestão clicada na lista '{suggestion_list_id}'.")
            except Exception as e:
                self.log(f"Erro ao clicar em sugestão: '{texto}', {e}.")

        except TimeoutException:
            self.log(f"Erro ao preencher ou selecionar a sugestão para o campo '{campo_id}'.")

    def _preencher_e_confirmar(self, campo, texto):
        """Insere texto em um campo e aguarda para pressionar Enter."""
        try:

            # Aguarda até que o campo esteja clicável
            campo = self.wait.until(EC.element_to_be_clickable(campo))
            
            # Enviar o texto
            campo.clear()  # Limpa o campo antes de inserir o texto
            campo.send_keys(texto)
            self.log(f"Texto '{texto}' inserido no campo com sucesso.")

            # Aguarda um pequeno intervalo para que o dropdown processe o texto
            self.wait.until(lambda driver: texto.lower() in campo.get_attribute("value").lower())
            self.log(f"Confirmação de que o campo contém o texto '{texto}'.")

            # Pausa breve
            sleep(0.5)
            
            # Pressionar Enter
            campo.send_keys(Keys.ENTER)
            self.log("Tecla Enter pressionada.")
        except TimeoutException:
            self.log(f"Erro ao inserir e confirmar o texto '{texto}'.")

    def _interagir_dropdown(self, dropdown_id, texto):
        """Interage com um dropdown customizado."""
        try:


            # Abre o dropdown
            dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
            dropdown.click()
            self.log(f"Dropdown '{dropdown_id}' aberto com sucesso.")

            # Localiza o campo de busca dentro do dropdown
            search_box = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f"#{dropdown_id} div.chosen-search input"))
            )

            sleep(0.5)
            
            # Preenche o texto e confirma com Enter
            self._preencher_e_confirmar(search_box, texto)
        except TimeoutException:
            self.log(f"Erro ao interagir com o dropdown '{dropdown_id}'.")

    def _preencher_campo_data_hora(self, campo_id, valor):
        """Preenche um campo de texto ou data/hora."""
        try:
            campo = self.wait.until(EC.element_to_be_clickable((By.ID, campo_id)))
            self.driver.execute_script("arguments[0].value = arguments[1];", campo, valor)
            self.log(f"Campo '{campo_id}' preenchido com valor '{valor}'.")
        except TimeoutException:
            self.log(f"Erro ao preencher o campo '{campo_id}'.")

    def fechar(self):
        """Finaliza o WebDriver."""
        try:
            input("aperte enter")
            self.driver.quit()
            self.log("WebDriver finalizado com sucesso.")
        except Exception as e:
            self.log(f"Erro ao finalizar o WebDriver: {e}")

    def salvar(self):
        """Salva o formulário."""
        try:
            salvar = self.wait.until(EC.element_to_be_clickable((By.ID, "idSubmit")))
            salvar.click()
            self.log("Formulário salvo com sucesso.")
        except TimeoutException:
            self.log("Erro ao salvar o formulário.")

    def incluir(self):
        """Clica no botão Incluir."""
        try:
            # Aguarda até que o botão esteja clicável
            botao_incluir = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.dt-button.btn_green"))
            )
            sleep(0.2)

            botao_incluir.click()
            self.log("Botão 'Incluir' clicado com sucesso.")
        except TimeoutException:
            self.log("Erro ao clicar no botão 'Incluir'. Verifique se ele está visível na página.")

    def finalizar(self):
        """Finaliza o ciclo clicando no botão 'finalizar'."""
        try:
            # Garantir que estamos no iframe central
            self._acessar_iframes_central()

            # Localizar e clicar no botão "Finalizar"
            botao_finalizar = self.wait.until(EC.element_to_be_clickable((By.ID, "idSubmit")))
            botao_finalizar.click()
            self.log("Botão 'Finalizar' clicado com sucesso.")

            # Aceitar o alerta automaticamente
            self.wait.until(EC.alert_is_present())  # Aguarda a presença do alerta
            self.driver.switch_to.alert.accept()
            self.log("Alerta aceito automaticamente.")

        except NoAlertPresentException:
            self.log("Nenhum alerta encontrado ao clicar em finalizar.")
        except Exception as e:
            self.log(f"Erro ao finalizar: {e}")

    def executar_planilha(self, file_path):
        """Executa as entradas da planilha com base em cabeçalhos e serviços."""
        try:
            # Carregar a planilha
            df = pd.read_excel(file_path)

            # Definir as colunas de cabeçalho
            colunas_cabecalho = [
                "obras_chosen", "contrato_chosen", "equipe_chosen",
                "tip_srv_chosen", "inputString", "cod_irr_chosen",
                "dat_srv", "hr_inic", "dat_srv2", "hr_fim"
            ]

            # Normalizar os dados da planilha (substituir NaN por None)
            df = df.where(pd.notnull(df), None)

            # Inicializar o cabeçalho anterior como vazio
            ultimo_cabecalho = []

                # Iterar pelas linhas da planilha
            try:
                for index, row in df.iterrows():
                    self.log(f"Processando linha {index}...")
                    # Extrair o cabeçalho atual
                    cabecalho_atual = {col: row[col] for col in colunas_cabecalho}

                    # Verificar se o cabeçalho mudou em relação ao último
                    if cabecalho_atual != ultimo_cabecalho:
                        # Se não for o primeiro ciclo, finalizar o ciclo anterior
                        if ultimo_cabecalho:
                            self.log(f"Ciclo: {index} para o cabeçalho anterior {index - 1}.")
                            self.finalizar()
                            self.log(f"Ciclo finalizado em {index} para o cabeçalho anterior {index - 1}.")

                        try:
                            self.preencher_cabecalho(row)

                            ultimo_cabecalho = cabecalho_atual

                            self.log(f"Novo cabeçalho processado {index}: {cabecalho_atual}")
                        except Exception as e:
                            self.log(f"Erro ao preencher o serviço {index}: {e}")

                    # Preencher o serviço correspondente à linha atual
                    self.preencher_servico(row)
                    self.log(f"Linha {index} processada com sucesso.")
            except Exception as e:
                self.log(f"Erro ao iterar pela planilha: {e}")

                # Finalizar o último ciclo após o loop
                self.finalizar()
                self.log("Último ciclo finalizado com sucesso.")

        except Exception as e:
            self.log(f"Erro ao processar a planilha: {e}")


# Execução do Script
if __name__ == "__main__":
    apontamento = Apontamento()
    try:
        apontamento.login() 
        apontamento.botao_serviço()
        apontamento._acessar_iframes_lateral()
        apontamento._acessar_iframes_central()
        apontamento.executar_planilha("dados_apontamento.xlsx")
    except Exception as e:
        apontamento.log(f"Erro inesperado: {e}")
    finally:
        apontamento.fechar()
        apontamento.log("Script finalizado.")