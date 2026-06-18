# 🚧 Automação SICRO – Extração, Consolidação e Auditoria de Orçamentos Rodoviários

![Python](https://img.shields.io/badge/Python-3.12-blue)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-green)
![Excel](https://img.shields.io/badge/Excel-Automation-darkgreen)
![SICRO](https://img.shields.io/badge/SICRO-Infraestrutura-orange).

Sistema desenvolvido em Python para automatizar a leitura, extração, consolidação e análise de relatórios SICRO, gerando uma planilha Excel estruturada e auditável para apoio à elaboração e verificação de orçamentos de infraestrutura.

---

## 📌 Visão Geral

A elaboração e análise de orçamentos rodoviários exige a manipulação de grandes volumes de informações provenientes de relatórios analíticos e sintéticos do SICRO.

Esse processo normalmente envolve:

- Leitura manual de centenas ou milhares de composições;
- Levantamento de mão de obra direta e indireta;
- Consolidação de equipamentos;
- Análise de materiais;
- Administração Local;
- Canteiro de Obras;
- Mobilização e Desmobilização;
- DMT, Tempo Fixo e Momento de Transporte;
- Conferências e auditorias técnicas.

O objetivo deste projeto é automatizar essas etapas, reduzindo tempo de processamento, minimizando erros operacionais e aumentando a rastreabilidade das informações.

---

## 🎯 Objetivo

Transformar relatórios SICRO em uma planilha Excel auditável, organizada e pronta para análise técnica pelo engenheiro orçamentista.

O sistema não substitui a análise de engenharia, mas automatiza tarefas repetitivas e operacionais, permitindo que o profissional foque na tomada de decisão.

---

## ⚙️ Principais Funcionalidades

### Extração de Dados

- Leitura automática de relatórios SICRO;
- Processamento de arquivos Analíticos e Sintéticos;
- Normalização de códigos de composição;
- Tratamento de inconsistências de entrada.

### Composições

- Identificação automática de composições principais;
- Extração de atividades auxiliares;
- Consolidação das informações por composição;
- Cálculo da produção das equipes.

### Mão de Obra

- Identificação dos blocos de mão de obra;
- Consolidação por função;
- Cálculo de efetivo;
- Resumo de colaboradores necessários para execução.

### Equipamentos

- Extração automática dos equipamentos utilizados;
- Consolidação das quantidades;
- Identificação de equipamentos produtivos e improdutivos.

### Materiais

- Extração dos materiais das composições;
- Consolidação para análise posterior.

### Transporte

- Extração de Tempo Fixo;
- Extração de Momento de Transporte;
- Consolidação de DMT.

### Administração Local

- Estruturação da Administração Local Variável;
- Resumo para apoio ao dimensionamento da equipe administrativa.

### Canteiro

- Estruturação do canteiro principal;
- Estruturação do canteiro complementar;
- Consolidação dos quantitativos.

### Auditoria

- Identificação de códigos não encontrados;
- Consolidação de inconsistências;
- Apoio à conferência técnica do orçamento.

---

## 📊 Estrutura do Arquivo Gerado

O sistema gera automaticamente uma planilha Excel contendo as seguintes abas:

| Aba | Finalidade |
|-------|-------|
| Base | Dados consolidados do processamento |
| Lista_CPU | Relação das composições encontradas |
| Sintético | Resumo dos serviços |
| BDI | Apoio ao cálculo do BDI |
| ABC | Curva ABC |
| Cronograma | Base para planejamento |
| Composição | Estrutura analítica das composições |
| ADM_Var | Administração Local Variável |
| ADM_Resumo | Resumo da Administração Local |
| CANT_Princ | Canteiro Principal |
| CANT_Complementar | Canteiro Complementar |
| CANT_Resumo | Resumo do Canteiro |
| DMT | Tempo Fixo e Momento de Transporte |
| Equipamentos | Relação consolidada de equipamentos |
| Pessoas | Consolidação de mão de obra |
| Não Encontrados | Códigos não localizados |
| Check_Orçamento | Verificações automáticas |

---

## 🖼️ Demonstração do Sistema

RESUMO: Indicadores gerais do orçamento processado.
<img width="1919" height="1023" alt="Resumo" src="https://github.com/user-attachments/assets/1cbc135b-c192-448c-9a2c-1d159cac28ed" />

SINTETICO: Serviços, quantitativos e custos consolidados.
<img width="1919" height="1020" alt="Sintetico" src="https://github.com/user-attachments/assets/8b6febac-f864-4d8e-88ae-5828beae2a8a" />

ANALITICO: Detalhamento completo das composições SICRO.
<img width="1915" height="1017" alt="Analitico" src="https://github.com/user-attachments/assets/b5d23f26-8508-4b44-b476-dcfb3a70c55c" />

DMT: Tempos fixos e momentos de transporte consolidados.
<img width="1917" height="1023" alt="DMT" src="https://github.com/user-attachments/assets/b9b75bcd-22b3-4d6f-936e-d95bb105bfe1" />

ADM: Recursos administrativos da obra consolidados.
<img width="1919" height="991" alt="ADM" src="https://github.com/user-attachments/assets/b4be48b8-1dc3-468f-a6e8-17ec8e8cf83e" />

CANTEIRO: Recursos de infraestrutura e apoio operacional da obra.
<img width="1919" height="1023" alt="Canteiro" src="https://github.com/user-attachments/assets/fdf7185b-35cd-4258-a22c-10514a9bb31e" />

---

## 🔄 Fluxo de Processamento

```text
Relatórios SICRO
        │
        ▼
Normalização dos Dados
        │
        ▼
Leitura das Composições
        │
        ▼
Extração dos Blocos
(A, B, C, D, E e F)
        │
        ▼
Consolidação
        │
        ├── Equipamentos
        ├── Mão de Obra
        ├── Materiais
        ├── DMT
        ├── Administração Local
        └── Canteiro
        │
        ▼
Geração do Excel
        │
        ▼
Validação e Auditoria
```

---

## 🛠️ Tecnologias Utilizadas

- Python
- Pandas
- OpenPyXL
- Regular Expressions (Regex)
- Excel
- Engenharia de Custos
- SICRO
- Metodologia DNIT

---

## 💡 Diferenciais do Projeto

Diferentemente de uma simples extração de dados, o projeto foi desenvolvido com foco em aplicações reais de Engenharia de Custos.

Os principais diferenciais incluem:

- Preservação da lógica utilizada por engenheiros orçamentistas;
- Estrutura de saída auditável;
- Consolidação automática de informações dispersas;
- Apoio ao levantamento de efetivo;
- Apoio ao dimensionamento de Administração Local;
- Apoio ao dimensionamento de Canteiro de Obras;
- Redução significativa de tempo de análise.

---

## 📈 Benefícios Obtidos

- Redução do tempo de processamento de orçamentos;
- Padronização das análises;
- Maior rastreabilidade dos dados;
- Menor risco de erros operacionais;
- Facilidade para auditoria e conferência;
- Apoio à tomada de decisão.

---

## 🚧 Limitações Atuais

- Algumas etapas ainda dependem de validação técnica do engenheiro;
- Mobilização e Desmobilização ainda possuem componentes orientados/manualizados;
- O sistema depende da estrutura padrão dos relatórios SICRO;
- Não substitui análise técnica especializada.

---

## 🔮 Roadmap

### Em desenvolvimento

- [ ] Histograma de Mão de Obra
- [ ] Evolução da Administração Local
- [ ] Evolução da aba Equipamentos
- [ ] Melhorias em Mobilização e Desmobilização
- [ ] Aprimoramento do Check_Orçamento
- [ ] Melhor tratamento de Atividades Auxiliares
- [ ] Interface Gráfica (GUI)
- [ ] Relatórios automáticos em PDF

### Futuras versões

- [ ] Integração com SINAPI
- [ ] Banco de dados de composições
- [ ] Dashboard em Power BI
- [ ] API para integração com outros sistemas
- [ ] Machine Learning para apoio à estimativa de recursos

---

## 📂 Estrutura do Projeto

```text
sicro-budget-automation/
│
├── main.py
├── requirements.txt
├── README.md
│
├── src/
│   ├── leitura_sicro.py
│   ├── extracao_composicoes.py
│   ├── processamento_mo.py
│   ├── processamento_equipamentos.py
│   ├── processamento_dmt.py
│   ├── gerar_excel.py
│   └── validacoes.py
│
├── templates/
│   └── modelo_excel.xlsx
│
├── docs/
│   ├── regras_calculo.md
│   ├── fluxo_sistema.md
│   └── imagens/
│
├── exemplos/
│   └── arquivo_saida.xlsx
│
└── tests/
```

---

## 👨‍💻 Autor

Lucas Melo

Engenheiro Civil | Engenheiro de Software

Atuação em:

- Engenharia de Custos
- Planejamento de Obras
- Fiscalização
- Automação de Processos
- Engenharia de Dados
- Desenvolvimento Python

---

## 📜 Licença

Este projeto foi desenvolvido para fins de estudo, pesquisa e aplicação profissional em Engenharia de Custos e Infraestrutura.

```
*"Automação aplicada à Engenharia de Custos, unindo Python, Excel e metodologia SICRO para reduzir retrabalho e aumentar a confiabilidade na análise de orçamentos rodoviários."*
```
