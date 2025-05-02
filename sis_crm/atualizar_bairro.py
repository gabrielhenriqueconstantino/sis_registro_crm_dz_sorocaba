"""
===========================================================================================
AUTOMAÇÃO DE BUSCA DE BAIRRO E PROBLEMA NO SITE CENTRAL 156 – PREFEITURA DE SOROCABA

Este script em Python automatiza o processo de atualização de registros de protocolos em 
no banco de dados SQLite com informações sobre *bairro* e *problema* associados a cada 
protocolo registrado no sistema de ouvidoria da Prefeitura de Sorocaba (Central 156).

Você deve usá-lo SOMENTE caso precise preencher muitos campos onde o valor de bairro e problema forem = None,
como no caso de você ter inserido muitas linhas vindas de uma planilha de excel, por exemplo, e não souber 
exatamente qual é o bairro e o assunto de de cada reclamação.

FUNCIONALIDADES PRINCIPAIS:
-------------------------------------------------------------------------------------------
1. **Acesso automatizado ao site da Prefeitura**:
   - Para cada protocolo sem bairro ou problema definidos no banco de dados, o script acessa
     a página correspondente no site `http://central156.sorocaba.sp.gov.br`.
   - O acesso é feito por meio do Selenium WebDriver com o navegador Google Chrome.

2. **Extração do bairro**:
   - A informação do bairro é obtida do bloco de endereço da reclamação.
   - É feito um tratamento de texto usando expressões regulares para isolar e limpar o nome do bairro.
   - Um caso especial é tratado: se o bairro for "Barão" ou "Barao", ele é padronizado como "Vila Barão".

3. **Extração do problema**:
   - O tipo de problema é extraído de um campo com múltiplos níveis de categoria (ex: "Reclamações / Saúde Pública / Animais Peçonhentos").
   - O script localiza o texto usando uma busca por labels com pelo menos duas barras ("/"), e captura a última parte (o nome do problema).

4. **Atualização no banco de dados**:
   - Após a extração, o script atualiza os campos `bairro` e `problema` no banco `sistema_protocolos.db`,
     Onde o valor for = None.
   - Cada atualização é feita individualmente com commit imediato para garantir persistência.

5. **Controle de execução e delays**:
   - O script imprime no console o andamento de cada protocolo.
   - Um pequeno `delay` (2 segundos) é adicionado entre requisições para evitar bloqueios por excesso de acessos ao site.

6. **Execução autônoma**:
   - O script pode ser executado diretamente ao clicar em "Run Python File" (modo script) e iniciará o processo automaticamente.

REQUISITOS:
-------------------------------------------------------------------------------------------
- Python 3 com as bibliotecas: `sqlite3`, `re`, `time`, `selenium`, `pathlib`.
- Google Chrome instalado e compatível com a versão do ChromeDriver especificada.
- Banco de dados SQLite com a tabela `protocolos`, contendo os campos `id`, `protocolo`, 
  `bairro`, e `problema`, entre outros.

OBSERVAÇÕES:
-------------------------------------------------------------------------------------------
- O caminho do ChromeDriver e do banco de dados pode ser ajustado conforme necessário.
- Se quiser rodar o navegador em segundo plano (sem abrir janela), mude `options.headless` para True.
- Use com cautela para não sobrecarregar o site da prefeitura com muitas requisições seguidas.

===========================================================================================
"""
import sqlite3
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path

# Caminho do banco de dados
DB_PATH = Path("database/db/sistema_protocolos.db")  # ajuste se estiver em outro lugar

# Caminho do ChromeDriver
CHROMEDRIVER_PATH = Path("chromedriver/chromedriver.exe")  # ajuste se necessário

def buscar_bairro_e_problema(protocolo):
    options = Options()
    options.headless = False  # Mude para True se quiser rodar em segundo plano
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    url = f"http://central156.sorocaba.sp.gov.br/atendimento/#/User/Request?protocolo={protocolo}&origem=2"
    driver.get(url)

    try:
        # === BUSCAR BAIRRO ===
        bloco_endereco = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "Endereço")]/parent::div'))
        )
        label_endereco = bloco_endereco.find_element(By.XPATH, './/label[@class="ng-binding"]')
        endereco_completo = label_endereco.text
        print(f"[{protocolo}] Endereço completo:", endereco_completo)

        bairro = None
        bairro_match = re.search(r'-\s*([^-/]+)\s*-\s*Sorocaba\s*/SP$', endereco_completo)
        if bairro_match:
            bairro = bairro_match.group(1).strip()
            bairro = re.sub(r'-', '', bairro).strip()
            if bairro.lower() in ["barão", "barao"]:
                bairro = "Vila Barão"

        # === BUSCAR PROBLEMA ===
        problema = None
        # A string que contém "Reclamações / Saúde Pública / Animais Peçonhentos" está dentro de um label genérico
        # Vamos buscar o primeiro label com esse formato
        labels = driver.find_elements(By.XPATH, '//label[@class="ng-binding"]')
        for label in labels:
            texto = label.text
            if texto.count(" / ") >= 2:
                partes = texto.split(" / ")
                problema = partes[-1].strip()
                print(f"[{protocolo}] Problema extraído: {problema}")
                break

        return (bairro if bairro else None), (problema if problema else None)

    except Exception as e:
        print(f"[{protocolo}] Erro ao buscar dados:", e)
        return None, None
    finally:
        driver.quit()


def atualizar_registros():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT id, protocolo FROM protocolos WHERE bairro IS NULL OR problema IS NULL")
    registros = cursor.fetchall()

    print(f"Total de registros para atualizar: {len(registros)}")

    for id_registro, protocolo in registros:
        print(f"\nAtualizando protocolo: {protocolo} (ID: {id_registro})")

        bairro, problema = buscar_bairro_e_problema(protocolo)
        if bairro:
            cursor.execute(
                "UPDATE protocolos SET bairro = ?, problema = ? WHERE id = ?",
                (bairro, problema, id_registro)
            )
            conn.commit()
            print(f"[{protocolo}] Atualizado: Bairro = {bairro}, Problema = {problema}")
        else:
            print(f"[{protocolo}] Bairro não encontrado, mantendo como None.")

        # Esperar um pouco para evitar bloqueio do site
        time.sleep(2)

    conn.close()
    print("\nAtualização concluída!")

if __name__ == "__main__":
    atualizar_registros()

