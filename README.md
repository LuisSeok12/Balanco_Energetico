# **Automação do Balanço Energetico com Python**

## **Descrição**
Este projeto é uma automação completa desenvolvida para integrar o sistema Thunders com planilhas Excel, visando a criação do balanço energetico usando uma planilha criada por mim mesmo como base. A aplicação combina automação web com manipulação avançada de arquivos Excel, tornando o processo mais eficiente e eliminando tarefas manuais repetitivas.

---

## **Tecnologias Utilizadas**

### **1. Python**
Linguagem principal utilizada para desenvolver o projeto, devido à sua versatilidade e vasta gama de bibliotecas.

### **2. Selenium**
- Usado para automação de navegadores.
- Realiza o login no sistema Thunders, navega entre páginas e baixa os arquivos necessários.
- Permite interagir com elementos dinâmicos e lidar com sites baseados em JavaScript.

### **3. Tkinter**
- Responsável pela interface gráfica.
- Captura dados do usuário, como matrícula, e-mail e senha, de forma amigável.
- Simplifica a interação inicial com o script.

### **4. Pandas**
- Biblioteca de manipulação de dados.
- Lê os arquivos Excel baixados e processa os dados em DataFrames.
- Realiza operações como filtragem e transformação dos dados.

### **5. OpenPyXL**
- Usada para editar e atualizar planilhas Excel.
- Escreve os dados processados em uma planilha consolidada, preservando fórmulas e formatos.

### **6. WebDriver Manager**
- Gerencia automaticamente o download do driver do navegador Chrome.
- Garante que o Selenium funcione sem configurações manuais complexas.

---

## **Funcionalidades**
1. **Automação Web com Selenium**:
   - Login automatizado no sistema Thunders.
   - Navegação entre páginas específicas.
   - Download automático de arquivos.

2. **Manipulação de Planilhas Excel**:
   - Leitura de arquivos Excel baixados.
   - Processamento e organização dos dados.
   - Escrita de informações em uma planilha consolidada, preservando fórmulas importantes.

3. **Interface Gráfica com Tkinter**:
   - Entrada de dados do usuário de maneira intuitiva.
   - Confirmação antes de executar o processo de automação.

---

## **Pré-requisitos**
- Python 3.8 ou superior.
- Dependências listadas no arquivo `requirements.txt`.

---

