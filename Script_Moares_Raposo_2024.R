################################################################################
# Título: Impacto da vacinação e variantes do SARS-CoV-2 nos desfechos graves da COVID-19 no Brasil, 2021-2022: estudo transversal
# Autores: Luiza Moraes, Letícia Raposo
# Instituição: Universidade Federal do Estado do Rio de Janeiro (UNIRIO)
# Data: Setembro de 2024
################################################################################

# Notas:
# - O script inclui o carregamento dos dados e seu tratamento inicial, seguido da aplicação de análises estatísticas
#   adequadas para cada variável de interesse.
# - Tabelas são geradas e salvas em formato .docx para facilitar a incorporação em relatórios e artigos.
#
# Licença:
# Este script é distribuído sob a licença CC-BY, permitindo sua reutilização com atribuição apropriada.
#
# Contato:
# - Letícia Raposo (leticia.raposo@uniriotec.br)
################################################################################

# Carregamento de pacotes necessários
library(readxl)     # Leitura de arquivos Excel
library(dplyr)      # Manipulação de dados
library(gtsummary)  # Criação de tabelas estatísticas
library(flextable)  # Exportação de tabelas para formatos de documentos

# 1. Leitura dos dados
dados <- read_excel("C:\\Users\\Leticia\\Google Drive\\UNIRIO\\Projetos\\2021\\Iniciação Científica - FAPERJ\\Luiza\\Artigo RL Vacinação\\Submissão RESS\\dados_artigo_Moraes_Raposo_2024.xlsx")

# 2. Tratamento dos dados
# Convertendo todas as variáveis, exceto as colunas especificadas (2, 28, 29), para o tipo fator.
dados[,-c(2,28,29)] <- dados[,-c(2,28,29)] %>% mutate_all(as.factor)

# 3. Configuração dos temas para as tabelas de resumo
theme_gtsummary_language(
  language = "pt",          # Define o idioma para Português
  decimal.mark = ",",       # Define a vírgula como separador decimal
  big.mark = ".",           # Define o ponto como separador de milhares
  iqr.sep = "-",            # Define o hífen como separador de intervalo interquartil
  ci.sep = "-",             # Define o hífen como separador para intervalos de confiança
  set_theme = TRUE          # Aplica essas configurações como tema padrão
)

# Função personalizada para formatação de porcentagens
list(
  "tbl_summary-fn:percent_fun" = function(x) sprintf(x * 100, fmt = '%#.1f')
) %>% set_gtsummary_theme()  # Aplica a função como tema padrão

# 4. Criação das Tabelas Resumidas

# 4.1 Tabela 1 - Resumo dos dados por variante do SARS-CoV-2
dados %>%
  tbl_summary(by = Variantes_SarsCoV2) %>%
  add_overall() %>%
  add_p() %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela1.docx")

# 4.2 Tabela 2 - Análise específica para variantes Delta e Ômicron

# 4.2.1 Delta
dados %>% 
  filter(Variantes_SarsCoV2 == "Delta") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19, -c(15:24)) %>%
  tbl_summary(by = Vacinacao_COVID_19_2) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 2)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela2_Delta.docx")

# 4.2.2 Ômicron
dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19, -c(15:24)) %>%
  tbl_summary(by = Vacinacao_COVID_19_2) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 2)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela2_Omicron.docx")

# 4.3 Tabela 3 - Comparação para a variante Ômicron
dados[,-c(15:24)] %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19_2) %>%
  tbl_summary(by = Vacinacao_COVID_19) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 2)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela3.docx")

# 5. Análise por variáveis específicas

# 5.1 Ajuste dos níveis da variável "Raça"
dados$Raca <- factor(dados$Raca, levels = c("Branca", "Parda/Preta", "Amarela/Indígena"))

# 5.2 Análise de Ventilação Mecânica Invasiva

# 5.2.1 Tabela descritiva para Ventilação Mecânica Invasiva
dados %>% 
  select(-Vacinacao_COVID_19_2, -UTI, -Obito_COVID_19) %>%
  tbl_summary(by = Ventilacao_Mecanica_Invasiva) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 2)) %>%
  as_flex_table()

# 5.2.2 Modelo de Regressão para Ventilação Mecânica Invasiva
tabMV1 <- dados %>%
  select(-Vacinacao_COVID_19_2, -UTI, -Obito_COVID_19) %>%
  tbl_uvregression(
    method = glm,
    y = Ventilacao_Mecanica_Invasiva,
    method.args = list(family = binomial),
    exponentiate = TRUE,
    pvalue_fun = ~ style_pvalue(.x, digits = 2)
  ) %>%
  bold_p() %>%
  bold_labels()

# Ajuste do modelo de regressão logística
dadosMV <- dados %>%
  select(-Vacinacao_COVID_19_2, -UTI, -Obito_COVID_19)

modMV <- glm(Ventilacao_Mecanica_Invasiva ~ ., data = dadosMV, family = binomial(link = "logit"))
modMV2 <- step(modMV, direction = "backward")
summary(modMV2)

tabmodMV2 <- modMV2 %>%
  tbl_regression(exponentiate = TRUE, pvalue_fun = ~ style_pvalue(.x, digits = 2)) %>%
  bold_p() %>%
  bold_labels()

# 5.3 Análise de Óbito por COVID-19

# 5.3.1 Tabela descritiva para Óbito por COVID-19
dados %>% 
  select(-Vacinacao_COVID_19_2) %>%
  tbl_summary(by = Obito_COVID_19) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 2)) %>%
  as_flex_table()

# 5.3.2 Modelo de Regressão para Óbito por COVID-19
tabDeath1 <- dados %>%
  select(-Vacinacao_COVID_19_2) %>%
  tbl_uvregression(
    method = glm,
    y = Obito_COVID_19,
    method.args = list(family = binomial),
    exponentiate = TRUE,
    pvalue_fun = ~ style_pvalue(.x, digits = 2)
  ) %>%
  bold_p() %>%
  bold_labels()

# Ajuste do modelo de regressão logística para óbito
dadosDeath <- dados %>%
  select(-Vacinacao_COVID_19_2)

modDeath <- glm(Obito_COVID_19 ~ ., data = dadosDeath, family = binomial(link = "logit"))
modDeath2 <- step(modDeath, direction = "backward")
summary(modDeath2)

tabmodDeath2 <- modDeath2 %>%
  tbl_regression(exponentiate = TRUE, pvalue_fun = ~ style_pvalue(.x, digits = 2)) %>%
  bold_p() %>%
  bold_labels()

# 6. Tabela Final Combinada para Ventilação Mecânica e Óbito
tbl_merge_MV <- tbl_merge(
  tbls = list(tabmodMV2, tabmodDeath2),
  tab_spanner = c("**Ventilação mecânica invasiva**", "**Óbito**")
)
tbl_merge_MV %>% as_flex_table() %>% save_as_docx(path = "Tabela4.docx")
