with open("README.md", "w", encoding="utf-8") as f:
    content = """# **README.md**  

# 🛠️ Robô de Apontamento Automático

Este projeto é um **robô automatizado** desenvolvido com **Selenium** e **Python** para realizar o preenchimento de formulários e serviços no sistema **Ceneged GPM**. A aplicação lê dados de uma planilha Excel e interage com os elementos na interface web, simulando ações humanas.

---

## 📋 **Funcionalidades**
- Realiza **login automático** no sistema.
- Preenche **cabeçalhos** e **serviços** com base nos dados fornecidos em uma planilha Excel.
- Identifica mudanças nos cabeçalhos e executa ações de finalização e reinício automaticamente.
- Interage com **dropdowns**, caixas de texto, e lida com sugestões automáticas.
- Gerencia **pop-ups** de alerta e aceita automaticamente.
- Gera **logs detalhados** de cada ação realizada.
- Permite execução em loop até o término de todos os registros da planilha.

---

## 🧪 **Tecnologias Utilizadas**
- **Python 3.12**
- **Selenium 4.x** (para automação de navegador)
- **pandas** (para manipulação de planilhas)
- **openpyxl** (para leitura e escrita de arquivos Excel)
- **chromedriver-autoinstaller** (opcional, para o gerenciamento do ChromeDriver)

---

## 📂 **Pré-requisitos**
1. **Python 3.12** instalado no sistema.
2. Google Chrome instalado (versão compatível com o ChromeDriver).
3. Ambiente virtual configurado.

### **Instalação das Dependências**
1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/robo-apontamento.git
   cd robo-apontamento
