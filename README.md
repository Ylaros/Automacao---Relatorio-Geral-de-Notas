# 📦 Automação – Relatório Geral de Notas

Este projeto representa uma automação completa e integrada que visa gerar, comparar e consolidar dados de notas fiscais eletrônicas a partir de múltiplas fontes. Ele utiliza recursos de VBA, Python + Selenium e relatórios do sistema contábil **Domínio (Thomson Reuters)** e do **SIEG** para gerar um relatório confiável e acionável sobre o status das notas fiscais do mês.

---

## 🎯 Objetivo

O objetivo principal é realizar a **comparação entre diferentes fontes de dados fiscais** (Domínio, SIEG e XMLs locais), garantindo:

1. Identificação de **notas não importadas**, **cancelamentos não refletidos** e **erros de integração** entre SIEG e Domínio;
2. Detecção de **notas anuladas via SEFAZ**;
3. Apresentação consolidada da **quantidade de documentos fiscais** e seus **valores totais por período**.

---

## 🧩 Tecnologias e Integrações

- **VBA (Excel)** – automação central para processamento, organização e geração dos relatórios finais;
- **Python + Selenium** – utilizado para login e download automático de documentos na plataforma SIEG;
- **SQL Anywhere (ODBC)** – para extração de relatórios personalizados do Domínio Contábil;
- **Aplicativo SIEG** – para geração de relatórios XML detalhados do mês;
- **Pastas de XML locais** – utilizadas como fonte complementar para validação.

---

## ⚙️ Pré-requisitos

Para executar esta automação, é necessário:

1. Baixar e instalar o aplicativo SIEG:  
   🔗 [Download do SIEG](https://d14tgtye96e903.cloudfront.net/Setup/InstalarSIEG_3.64.zip)

2. Ter o ambiente Python configurado com Selenium

3. Ter acesso ao gerador de relatórios do sistema Domínio (via ODBC / SQL Anywhere)

---

## 📌 Fluxo resumido da automação

1. **Executar a automação VBA** e definir o período a ser analisado.
2. **Gerar os relatórios do SIEG** com detalhamento de produtos e CFOPs.
3. **Extrair relatórios do Domínio**:
   - Cupons Fiscais
   - Entradas
   - Saídas
   - Empresas ativas
4. **Executar o relatório base em Excel (`.xlsm`)**, consolidando os dados.
5. **Gerar arquivo final consolidado** com o nome "Relatório Geral de Notas - [período]".
6. **Organizar os arquivos gerados** em estrutura de pastas por ano/mês/dia.

---

## ❗ Importante

> ⚠️ Este repositório **não contém dados reais nem arquivos vinculados a clientes**. Todos os dados são fictícios ou genéricos para fins de demonstração.

---

📷 Capturas de Tela
|Exemplo 1 | [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/blob/main/Imagens/Relat%C3%B3rio%20Geral%20de%20Notas%2001.png]
|Exemplo 2 | [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/blob/main/Imagens/Relat%C3%B3rio%20Geral%20de%20Notas%2002.png]
|Exemplo 3 | [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/blob/main/Imagens/Relat%C3%B3rio%20Geral%20de%20Notas%2003.png]
|Exemplo 4 | [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/blob/main/Imagens/Relat%C3%B3rio%20Geral%20de%20Notas%2004.png]
|Exemplo 5 | [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/blob/main/Imagens/Relat%C3%B3rio%20Geral%20de%20Notas%2005.png]

📑 Automações Envolvidas
🔹 Python + Selenium - [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/tree/main/Python%20Exec]

🔹 Códigos VBA - [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/tree/main/VBA%20Codes]

🔹 Códigos SQL Anywhere - [https://github.com/Ylaros/Automacao-Relatorio-Geral-de-Notas/tree/main/SQL%20Anywhere]


## ✍️ Autor

**Aloyr Rezende**  
🔗 [LinkedIn](https://www.linkedin.com/in/aloyr-rezende)

---

## 💬 Licença

Este projeto é de uso interno e educacional. Caso deseje adaptar ou reutilizar partes da automação, sinta-se à vontade para contribuir ou propor melhorias.



