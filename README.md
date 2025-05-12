# 🩺📊 SisconProtDZ - Sistema de Consulta, Registro e Análise de Protocolos de Reclamação
---

**SisconProtDZ** é uma aplicação desktop para registro, controle e análise de protocolos oriundos da Ouvidoria do Munícipio de Sorocaba, onde os cidadãos podem fazer pedidos e solicitações para a prefeitura - incluindo visitas para vigilância em saúde pública, papel desempenhado pela Divisão de Zoonoses. Desenvolvido com Python e interface gráfica usando `Tkinter` e `CustomTkinter`, o sistema visa facilitar o gerenciamento de solicitações como visitas técnicas, denúncias de focos de dengue e demais demandas da população.

---

## 🧰 Funcionalidades

* 📝 Registro e edição de protocolos
* 🔍 Busca automática de bairro pelo número de protocolo (via site da prefeitura)
* 📊 Geração de gráficos estatísticos (pizza e barra)
* 🗃️ Backup e restauração de dados
* 🔎 Filtros por data, bairro, situação, problema e área
* 📋 Edição direta dos dados na tabela, como em uma planilha
* 🌗 Suporte a tema claro e escuro
* 🖥️ Interface adaptada para resoluções 1024x768

---

## 📂 Estrutura de Pastas

```
sis_crm/
├── atualizar_bairro/          # Funções de busca de bairro (Selenium)
├── chromedriver/              # Driver do navegador Chrome (para Selenium)
├── database/
│   ├── backup/                # Arquivos de backup
│   ├── db/                    # Banco de dados SQLite
│   └── insert_sql/            # Scripts de inserção
├── font/                      # Fontes personalizadas (se houver)
└── img/
    ├── icons/
    │   ├── icons_app/         # Ícones da aplicação
    │   └── icons_window/      # Ícones da janela
    └── logo/                  # Logos do sistema
```

---

## 💻 Instalação e uso

### 🔧 Requisitos

* **Sistema operacional:** Windows 10 ou superior
* **Conexão com internet** (para funcionalidades que usam o site da prefeitura)

### 🚀 Executar o sistema

1. Faça o download da última [release](https://github.com/gabrielhenriqueconstantino/sis_registro_crm_dz_sorocaba/releases/tag/v1.0.0).
2. Extraia o conteúdo (caso esteja zipado).
3. Execute o arquivo `SisconProtDZ.exe`.
4. O banco de dados será criado automaticamente na primeira execução, caso ainda não exista.

---

## 📈 Análises e Gráficos

* O botão “Análise” permite gerar gráficos por:

  * **Área administrativa**
  * **Tipo de problema**
  * Agrupamento por **quantidade** ou **tipo**

---

## 🛠️ Tecnologias Utilizadas

* Python 3.13
* Tkinter / CustomTkinter
* SQLite
* Selenium (busca de bairros via protocolo)
* PyInstaller (geração do executável)

---

## 📦 Como compilar (para desenvolvedores)

```bash
pyinstaller --onefile --noconsole
--add-data "img\\icons\\icons_app\\;img\\icons\\icons_app"
--add-data "img\\icons\\icons_window\\;img\\icons\\icons_window"
--add-data "img\\logo\\*;img\\logo"
--icon=img\\icons\\icons_window\\logo_sis_crm.ico
--name "SisconProtDZ" app.py
```

---

## 📄 Licença

Este projeto é livre e pode ser utilizado e modificado por setores da de uso interno da Prefeitura Municipal de Sorocaba

---
