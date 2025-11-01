
import os
import time
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException


URL_LOGIN = "https://www.cenprotsc.com.br/ieptb/view/ferramenta/login.xhtml"
URL_CNPJ = "https://www.cenprotsc.com.br/ieptb/view/ferramenta/buscacnpj.xhtml"
USUARIO = ""
SENHA = ""
ARQUIVO_SAIDA = "resultado_cnpjs.xlsx"
ARQUIVO_ENTRADA = r"Gui - Protestos/CNPJconsulta.txt"


EMAIL_REMETENTE = ""
SENHA_EMAIL = ""
EMAIL_DESTINO = ""
ASSUNTO_EMAIL = "Relatório de CNPJs - Consulta Automática"
MENSAGEM_EMAIL = """
Olá,

Segue em anexo o relatório com os resultados das consultas de CNPJs.

Atenciosamente,
Automação Python by Guilherme Morais and Gabriel Lopes 🤖
"""


def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    return driver

def fazer_login(driver):
    driver.get(URL_LOGIN)
    try:
        usuario = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Usuário']"))
        )
        usuario.clear()
        usuario.send_keys(USUARIO)

        senha = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Senha']"))
        )
        senha.clear()
        senha.send_keys(SENHA)

        btn_login = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "btnLogin"))
        )
        btn_login.click()

        time.sleep(2)
        # Se houver iframe após login
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            driver.switch_to.frame(iframes[0])
        print("✅ Login realizado com sucesso!")
        return True
    except Exception as e:
        print(f"❌ Falha no login: {e}")
        return False


def consultar_cnpj(driver, cnpj):
    try:
        print(f"🔎 Consultando {cnpj}...")

        driver.get(URL_CNPJ)
        time.sleep(2)  
        campo_cnpj = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.NAME, "j_idt1006"))
        )
        campo_cnpj.clear()
        campo_cnpj.send_keys(cnpj)
        campo_cnpj.send_keys(Keys.TAB)
        time.sleep(1)

        btn_buscar = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "btnConsultarCNPJ"))
        )
        driver.execute_script("arguments[0].click();", btn_buscar)

       
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//label[text()='Nome:']/following-sibling::div/span")
            )
        )

        
        def pegar_valor(label_texto):
            try:
                elemento = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, f"//label[text()='{label_texto}']/following-sibling::div/span")
                    )
                )
                return elemento.text.strip()
            except:
                return ""

        resultado = {
            "CNPJ": cnpj,
            "Nome": pegar_valor("Nome:"),
            "Nome Fantasia": pegar_valor("Nome Fantasia:"),
            "Email": pegar_valor("E-mail:"),
            "Telefone": pegar_valor("Fone:"),
            "Responsável": pegar_valor("Responsável:"),
            "Cidade": pegar_valor("Cidade:"),
            "Atividade Principal": pegar_valor("AtividadePrincipal:"),
        }

        print(f"📌 {resultado}")
        return resultado

    except Exception as e:
        print(f"⚠️ Erro ou nenhum resultado para {cnpj}: {e}")
        return {
            "CNPJ": cnpj,
            "Nome": "",
            "Nome Fantasia": "",
            "Email": "",
            "Telefone": "",
            "Responsável": "",
            "Cidade": "",
            "Atividade Principal": "",
        }

# ========================
# FUNÇÕES DE ARQUIVOS
# ========================
def carregar_cnpjs():
    if ARQUIVO_ENTRADA.endswith(".xlsx"):
        df = pd.read_excel(ARQUIVO_ENTRADA)
        return df.iloc[:, 0].astype(str).tolist()
    elif ARQUIVO_ENTRADA.endswith(".txt"):
        with open(ARQUIVO_ENTRADA, "r", encoding="utf-8") as f:
            return [linha.strip() for linha in f if linha.strip()]
    else:
        raise ValueError("Formato de arquivo não suportado.")

def carregar_resultados_existentes():
    if os.path.exists(ARQUIVO_SAIDA):
        df = pd.read_excel(ARQUIVO_SAIDA)
        return df.to_dict("records")
    return []

def enviar_email_com_anexo():
    print("📨 Enviando e-mail com o resultado...")
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = EMAIL_DESTINO
    msg["Subject"] = ASSUNTO_EMAIL
    msg.attach(MIMEText(MENSAGEM_EMAIL, "plain"))

    with open(ARQUIVO_SAIDA, "rb") as f:
        anexo = MIMEApplication(f.read(), Name=os.path.basename(ARQUIVO_SAIDA))
    anexo["Content-Disposition"] = f'attachment; filename="{os.path.basename(ARQUIVO_SAIDA)}"'
    msg.attach(anexo)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_EMAIL)
            smtp.send_message(msg)
        print("✅ E-mail enviado com sucesso!")
    except Exception as e:
        print(f"❌ Falha ao enviar e-mail: {e}")


def main():
    todos_resultados = carregar_resultados_existentes()
    cnpjs_realizados = {r["CNPJ"] for r in todos_resultados}
    cnpjs = carregar_cnpjs()
    pendentes = [c for c in cnpjs if c not in cnpjs_realizados]

    print(f"🔁 {len(pendentes)} CNPJs pendentes para consultar.")

    driver = None
    tentativa = 0

    while pendentes:
        try:
            if driver is None:
                driver = iniciar_driver()
                if not fazer_login(driver):
                    raise Exception("Falha no login.")

            total = len(pendentes)
            for i, cnpj in enumerate(pendentes[:], start=1):
                print(f"\n📊 Progresso: {i}/{total} | Restantes: {total - i}")
                resultado = consultar_cnpj(driver, cnpj)
                todos_resultados.append(resultado)
                pd.DataFrame(todos_resultados).to_excel(ARQUIVO_SAIDA, index=False)
                pendentes.remove(cnpj)
                time.sleep(3)  
        except (WebDriverException, TimeoutException) as e:
            print(f"⚠️ Erro no navegador: {e}")
            tentativa += 1
            if driver:
                driver.quit()
                driver = None
            print(f"⏳ Reiniciando navegador (tentativa {tentativa})...")
            time.sleep(5)

        except Exception as e:
            print(f"❌ Erro inesperado: {e}")
            if driver:
                driver.quit()
            break

    if driver:
        driver.quit()

    print(f"✅ Processo finalizado! Total consultado: {len(todos_resultados)}")
    enviar_email_com_anexo()


if __name__ == "__main__":
    main()
CONSULTA_CADASTRO.txt
Exibindo CONSULTA_CADASTRO.txt.
