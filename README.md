# Sistema de Simulação e Gerenciamento de Consórcios

Este repositório contém o código-fonte e a documentação do **Trabalho de Conclusão de Curso (TCC)** desenvolvido por **Pericles Ferreira Martins**.

O projeto consiste em uma solução híbrida para automação de processos em consórcios, integrando a capacidade de processamento de dados do **Python** com a interface gerencial do **Microsoft Excel (VBA)**.

![Status do Projeto](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Python](https://img.shields.io/badge/Python-3.13.7-blue)
![Excel](https://img.shields.io/badge/Excel-VBA-green)

![exemploUso](imagens/exemplo_de_uso.gif)

## 🎯 Objetivo do Projeto

Resolver o problema da morosidade e da suscetibilidade a erros no processo manual de leitura de grupos de consórcio. O sistema automatiza a extração de dados de relatórios em PDF e alimenta um simulador financeiro interativo.

## 🛠️ Arquitetura do Sistema

O sistema opera em duas camadas principais:

1.  **Backend (ETL - Extract, Transform, Load):**
    - Desenvolvido em **Python**.
    - Responsável por varrer uma pasta, ler arquivos `.pdf`, tratar inconsistências de formatação e exportar uma base de dados consolidada em Excel (`.xlsx`).
2.  **Frontend (Interface e Simulação):**
    - Desenvolvido em **Excel (VBA)**.
    - Interface onde o usuário realiza as simulações de lances, visualiza dashboards e gera propostas. O VBA consome a base de dados gerada pelo Python.

## 🚀 Tecnologias Utilizadas

- **Linguagem:** Python 3.13.7
- **Bibliotecas Python:**
  - `pdfplumber`: Extração de tabelas em PDFs.
  - `pandas`: Manipulação, limpeza e estruturação de dados.
  - `numpy`: Cálculos e tratamento de dados numéricos.
  - `openpyxl`: Engine para gravação de arquivos Excel.
  - `os`: Manipulação de sistema de arquivos.
- **Plataforma:** Microsoft Excel (.xlsm)
  - **VBA**: Automação de formulários, lógica de simulação financeira e controle de interface (Zoom/Impressão).

## 📘 Documentação Técnica dos Módulos

Abaixo detalhes do funcionamento lógico dos principais scripts que compõem o sistema.

### 📊 Frontend: VBA (Excel)

A lógica de negócios e a interface do usuário foram construídas através de módulos VBA para facilitar a manutenção.

#### 1. Módulo Principal (`Sub Simulacao`)

É o "coração" do sistema.

- **Entrada:** Coleta inputs do usuário via `InputBox` (Número do Grupo e % de Lance).
- **Processamento:** Realiza uma busca na base de dados importada. Ao encontrar o grupo correspondente, executa cálculos financeiros (cálculo de nova parcela após lance embutido/livre e projeção de prazo restante).
- **Saída:** Preenche dinamicamente a planilha "Simular" com os resultados e totais acumulados.

#### 2. Módulo de Impressão (`Sub imprimir`)

Gerencia a saída do relatório final, adaptando-se à necessidade do usuário:

- **Modo Físico (Opção 1):** Seleciona layouts otimizados para impressoras (tons de cinza/alto contraste) para facilitar a visualização das barras.
- **Modo Digital (Opção 2):** Seleciona layouts coloridos ideais para exportação em PDF e envio via WhatsApp.
- **Lógica:** O código verifica quantos grupos foram simulados (1, 2 ou 3) para escolher a aba de impressão correta, evitando gráficos em branco.

#### 3. Módulo de Interface/Zoom (`Sub AjustarZoomPlanilhas`)

Garante a responsividade da aplicação.

- O script identifica a resolução da tela do usuário e aplica um `ActiveWindow.Zoom` baseado em uma seleção de células (`Range`).
- Isso assegura que o Dashboard e os botões de comando estejam sempre visíveis e centralizados, independentemente se o monitor é 13" ou 24".

#### 4. Inicialização (`Workbook_Open`)

Evento disparado automaticamente ao abrir o arquivo.

- Prepara o ambiente de trabalho, definindo variáveis globais e, opcionalmente, travando a área de rolagem (`ScrollArea`) para criar uma experiência de "sistema", impedindo que o usuário final acesse áreas de rascunho da planilha.

## 📝 Guia das Planilhas:

O arquivo Excel (Simulador_Consórcio.xlsm) é composto por diversas abas, divididas entre Interface do Usuário (Frontend), Banco de Dados (Backend) e Layouts de Impressão. Abaixo a descrição de cada uma:

1. Interface do Usuário

   - **dashboard:** Painel visual com gráficos de coluna, informando quantidade de lances e contemplações por grupo. É a tela inicial do sistema.

   - **simular:** A tela principal de operação. É aqui que o usuário insere o número do grupo e o percentual de lance para receber os cálculos de parcelas, prazos...

   - **sobre:** Contém instruções breves de uso.

2. Dados e Processamento

   - **BaseDados:** O coração do sistema. É aqui que os dados tratados pelo Python devem ser colados.

   - **BASE:** Planilha serve para armazenar informações de lances (MÍN, MÁX, MED) e contemplações por grupo (QTD).

   - **DadosGrafico:** Aba técnica retorna e formata os grupos informados na aba "simular" para alimentar os gráficos de impressão.

3. Relatórios e Saída

   - impressaoFolha1GColor / 2G / 3G: Layouts pré-formatados para exportação em PDF ou impressão física, limitados a até 3 gráficos correspondentes aos primeiros grupos informados.

   - O VBA seleciona automaticamente qual dessas abas exibir baseando na quantidade de grupos simulados (1, 2 ou 3 grupos), garantindo que o relatório final não tenha gráficos com espaços em branco.

## 🚀 Como Executar o Projeto

Siga as etapas abaixo para configurar o ambiente e realizar uma simulação.

**⚠️ Configuração Inicial Necessária:**
Antes de executar o script pela primeira vez, é necessário ajustar o caminho da pasta base para o seu ambiente local:

1. Abra o arquivo `ETL.py` em um editor de texto ou IDE.
2. Localize a variável `base_path` (linha 115).
3. Altere o caminho para o diretório onde você salvou a pasta do projeto no seu computador.
   - Exemplo: De `C:\Users\{usuario}\OneDrive\...` para `C:\Users\{usuario}\Documents\Simulador-Consorcio`.

### Passo 1: Processamento de Dados (Pule se quiser usar a base de dados atual)

1.  Insira os arquivos **.PDF** (extratos dos grupos) dentro da pasta `PDF/` que está na raiz do projeto.
2.  No terminal, execute o script de automação:
    ```bash
    python ETL.py
    ```
3.  Aguarde a mensagem de conclusão. O script irá gerar/atualizar o arquivo `tabelas_banco.xlsx` dentro de `PDF/XLSX/`.
4.  **Atualização da Base:** Abra o arquivo gerado (`tabelas_banco.xlsx`) e copie o conteúdo. No arquivo do Simulador, cole os dados na aba de Base de Dados, substituindo **apenas as colunas de cor cinza**.
    - Nota: As colunas de cor **roxa** possuem fórmulas automáticas e não devem ser alteradas.

### Passo 2: Utilizando o Simulador

1.  Abra o arquivo `Simulador_Consórcio.xlsm`.
2.  ⚠️ **Importante:** Ao abrir, o Excel solicitará permissão para executar scripts. Clique em **"Habilitar Conteúdo"**. Sem isso, os botões e automações não funcionarão.
3.  Navegue até a aba **Simular**.
4.  Clique no botão de simulação e insira o **número do grupo** e o **percentual de lance** desejado conforme os dados extraídos.

Pronto! Obterá uma simulação completa em questão de segundos.

### 🔐 Senhas e Acesso

Para facilitar a avaliação e os testes, todas as proteções do sistema foram configuradas com uma senha padrão.

- **Senha Padrão:** `123`
- **Onde é solicitada:** Desbloqueio de planilhas e acesso ao código fonte VBA (Alt+F11).

> ⚠️ **Nota:** Caso este sistema venha a ser implementado em um ambiente real de produção, recomendo fortemente a alteração dessas senhas para garantir a integridade dos dados e do código.

---

## 👨‍💻 Autor

<a href="https://github.com/PericlesFMartins">
 <img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/189674643?v=4" width="100px;" alt=""/>
 <br />
 <sub><b>Pericles Ferreira Martins</b></sub>
</a>

Este projeto foi desenvolvido como parte do **Trabalho de Conclusão de Curso (TCC)** para o curso de Engenharia de software, em Concórdia - SC 03/2025.

O objetivo foi unir conhecimentos de **Engenharia de Dados (Python)** e **Automação (VBA)** para resolver um problema real de negócio.

[![Linkedin Badge](https://img.shields.io/badge/-LinkedIn-blue?style=flat-square&logo=Linkedin&logoColor=white)](https://www.linkedin.com/in/pericles-ferreira-martins-475b8114a/)
[![Gmail Badge](https://img.shields.io/badge/-Gmail-c14438?style=flat-square&logo=Gmail&logoColor=white&link=mailto:periclesrbyamartins@gmail.com)](mailto:periclesrbyamartins@gmail.com)

---

Desenvolvido com muito café.
