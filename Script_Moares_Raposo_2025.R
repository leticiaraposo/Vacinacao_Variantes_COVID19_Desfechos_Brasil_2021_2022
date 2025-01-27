################################################################################
# Título: Impacto da vacinação e variantes do SARS-CoV-2 nos desfechos graves da COVID-19 no Brasil, 2021-2022: estudo transversal
# Autores: Luiza Moraes, Letícia Raposo
# Instituição: Universidade Federal do Estado do Rio de Janeiro (UNIRIO)
# Data: Janeiro de 2025
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
library(sandwich)   # Fornece estimadores de variância robustos para modelos de regressão
library(lmtest)     # Complementa análises de modelos lineares ao oferecer testes estatísticos específicos

# 1. Leitura dos dados
dados <- read_excel("dados_artigo_Moraes_Raposo_2024.xlsx")

# 2. Tratamento dos dados
# Convertendo todas as variáveis, exceto as colunas especificadas (2, 14, 15), para o tipo fator.
dados[,-c(2,14,15)] <- dados[,-c(2,14,15)] %>% mutate_all(as.factor)

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

# Avaliando a normalidade por grupo
# Carregar os pacotes necessários
library(dplyr)  # Para manipulação de dados

# Avaliar a normalidade por grupos da variável "Variantes_SarsCoV2"
dados %>%
  group_by(Variantes_SarsCoV2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Idade, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>%
  group_by(Variantes_SarsCoV2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Duracao_Hospitalizacao, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>%
  group_by(Variantes_SarsCoV2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Tempo_Sintomas_Hospital, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

# 4.1 Tabela 1 - Resumo dos dados por variante do SARS-CoV-2
dados %>%
  tbl_summary(by = Variantes_SarsCoV2) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela1.docx")

# Realização de testes de Dunn para diferenças par a par entre grupos
library(DescTools)

DunnTest(dados$Idade, dados$Variantes_SarsCoV2) 
DunnTest(dados$Duracao_Hospitalizacao, dados$Variantes_SarsCoV2)
DunnTest(dados$Tempo_Sintomas_Hospital, dados$Variantes_SarsCoV2)

# 4.2 Tabela 2 - Análise específica para variantes Delta e Ômicron

# 4.2.1 Delta

# Avaliar a normalidade por grupos da variável "Vacinacao_COVID_19_2"
dados %>% 
  filter(Variantes_SarsCoV2 == "Delta") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Idade, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Delta") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Duracao_Hospitalizacao, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Delta") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Tempo_Sintomas_Hospital, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Delta") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19, -c(7:10)) %>%
  tbl_summary(by = Vacinacao_COVID_19_2) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela2_Delta.docx")

# 4.2.2 Ômicron

# Avaliar a normalidade por grupos da variável "Vacinacao_COVID_19_2"
dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Idade, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Duracao_Hospitalizacao, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19_2) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Tempo_Sintomas_Hospital, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19, -c(7:10)) %>%
  tbl_summary(by = Vacinacao_COVID_19_2) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela2_Omicron.docx")

# 4.3 Tabela 3 - Comparação para a variante Ômicron

# Avaliar a normalidade por grupos da variável "Vacinacao_COVID_19_2"
dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Idade, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Duracao_Hospitalizacao, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  group_by(Vacinacao_COVID_19) %>%  # Agrupa os dados por grupo
  summarise(
    p_value = ks.test(Tempo_Sintomas_Hospital, "pnorm")$p.value  # Executa o teste de Kolmogorov-Smirnov para cada grupo
  ) %>%
  mutate(Resultado = ifelse(p_value < 0.05, "Rejeita normalidade", "Não rejeita normalidade"))

dados[,-c(7:10)] %>% 
  filter(Variantes_SarsCoV2 == "Ômicron") %>%
  select(-Variantes_SarsCoV2, -Vacinacao_COVID_19_2) %>%
  tbl_summary(by = Vacinacao_COVID_19) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table() %>%
  save_as_docx(path = "Tabela3.docx")

dados_omicron <- dados %>%
  filter(Variantes_SarsCoV2 == "Ômicron")

DunnTest(dados_omicron$Idade, dados_omicron$Vacinacao_COVID_19) 
DunnTest(dados_omicron$Duracao_Hospitalizacao, dados_omicron$Vacinacao_COVID_19)
DunnTest(dados_omicron$Tempo_Sintomas_Hospital, dados_omicron$Vacinacao_COVID_19)

# 5. Análise por variáveis específicas

# 5.1 Ajuste dos níveis da variável "Raça" e ordenação dos níveis de Variante e Vacinação
dados$Raca <- factor(dados$Raca, levels = c("Branca", "Parda/Preta", "Amarela/Indígena"))

dados$Variantes_SarsCoV2 <- factor(dados$Variantes_SarsCoV2, levels = c("Gama",
                                                                        "Delta",
                                                                        "Ômicron"))

dados$Vacinacao_COVID_19 <- factor(dados$Vacinacao_COVID_19, levels = c("Não vacinado",
                                                                        "Incompleta",
                                                                        "Primária",
                                                                        "Completa"))
# 5.2 Análise de Ventilação Mecânica Invasiva

# 5.2.1 Tabela descritiva para Ventilação Mecânica Invasiva
dados %>% 
  select(-Vacinacao_COVID_19_2, -UTI, -Obito_COVID_19) %>%
  tbl_summary(by = Ventilacao_Mecanica_Invasiva) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table()

# 5.2.2 Modelo de Regressão de Poisson com Variância Robusta para Ventilação Mecânica Invasiva

# Criar uma cópia dos dados originais para manipulação
dadosp <- dados

# Converter a variável "Ventilacao_Mecanica_Invasiva" para numérica e ajustar a escala para iniciar em 0
# (subtraindo 1, assumindo que a variável original está codificada como 1 e 2)
dadosp$Ventilacao_Mecanica_Invasiva <- as.numeric(dadosp$Ventilacao_Mecanica_Invasiva) - 1

# Remover variáveis desnecessárias do conjunto de dados para o modelo de Poisson
# (aqui estão sendo excluídas as variáveis "Vacinacao_COVID_19_2", "UTI" e "Obito_COVID_19")
dadosMV <- dadosp %>%
  select(-Vacinacao_COVID_19_2, -UTI, -Obito_COVID_19)

# Ajustar um modelo de regressão de Poisson com link log para a variável resposta
# "Ventilacao_Mecanica_Invasiva" e todas as demais variáveis do conjunto de dados como preditoras
modMV <- glm(Ventilacao_Mecanica_Invasiva ~ ., data = dadosMV, 
             family = poisson(link = "log"))

# Ajustar a variância robusta do modelo para corrigir possíveis violações de pressupostos
# (como superdispersão) utilizando a matriz de variância-covariância robusta (HC0)
modMV_robust <- coeftest(modMV, vcov = vcovHC(modMV, type = "HC0"))

# Realizar seleção de variáveis pelo método "backward" (eliminação retrógrada)
# para identificar o modelo mais parcimonioso
modMV2 <- step(modMV, direction = "backward")

# Ajustar a variância robusta para o modelo selecionado
modMV2_robust <- coeftest(modMV2, vcov = vcovHC(modMV2, type = "HC0"))

# Calcular a matriz de variância-covariância robusta para o modelo selecionado
robust_se <- vcovHC(modMV2, type = "HC0")

# Obter os coeficientes robustos do modelo final
robust_results <- coeftest(modMV2, vcov = robust_se)

# Extrair os coeficientes estimados do modelo final
coef_estimates <- coef(modMV2)

# Calcular os erros-padrão robustos a partir da matriz de variância-covariância
robust_se_values <- sqrt(diag(robust_se))

# Calcular as razões de prevalência (exponenciação dos coeficientes do modelo)
PR <- exp(coef_estimates)

# Calcular os intervalos de confiança de 95% para as razões de prevalência
lower_CI <- exp(coef_estimates - 1.96 * robust_se_values)  # Limite inferior
upper_CI <- exp(coef_estimates + 1.96 * robust_se_values)  # Limite superior

# Criar uma tabela final com as variáveis, razões de prevalência e seus intervalos de confiança
results_table <- data.frame(
  Variable = names(coef_estimates),        # Nomes das variáveis
  Prevalence_Ratio = round(PR, 2),        # Razão de prevalência (RP) arredondada
  Lower_95CI = round(lower_CI, 2),        # Limite inferior do IC95%
  Upper_95CI = round(upper_CI, 2)         # Limite superior do IC95%
)

# Teste de deviance
anova(modMV2, test = "Chisq")

# Cálculo do pseudo R²
1 - (modMV2$deviance / modMV2$null.deviance) #6,80%

# 5.3 Análise de Óbito por COVID-19

# 5.3.1 Tabela descritiva para Óbito por COVID-19
dados %>% 
  select(-Vacinacao_COVID_19_2) %>%
  tbl_summary(by = Obito_COVID_19) %>%
  add_overall() %>%
  add_p(pvalue_fun = function(x) style_pvalue(x, digits = 3)) %>%
  as_flex_table()

# 5.3.2 Modelo de Regressão de Poisson robusta para Óbito por COVID-19
dadosp$Obito_COVID_19 <- as.numeric(dadosp$Obito_COVID_19) - 1

dadosDeath <- dadosp %>%
  select(-Vacinacao_COVID_19_2)

# Modelo Poisson
modDeath <- glm(Obito_COVID_19 ~ ., data = dadosDeath, 
                family = poisson(link = "log"))

# Ajustar a variância robusta do modelo para corrigir possíveis violações de pressupostos
# (como superdispersão) utilizando a matriz de variância-covariância robusta (HC0)
modDeath_robust <- coeftest(modDeath, vcov = vcovHC(modDeath, type = "HC0"))

# Realizar seleção de variáveis pelo método "backward" (eliminação retrógrada)
# para identificar o modelo mais parcimonioso
modDeath2 <- step(modDeath, direction = "backward")

# Ajustar a variância robusta para o modelo selecionado
modDeath2_robust <- coeftest(modDeath2, vcov = vcovHC(modDeath2, type = "HC0"))

# Calcular a matriz de variância-covariância robusta para o modelo selecionado
robust_se2 <- vcovHC(modDeath2, type = "HC0")

# Obter os coeficientes robustos do modelo final
robust_results2 <- coeftest(modDeath2, vcov = robust_se2)

# Extrair os coeficientes estimados do modelo final
coef_estimates2 <- coef(modDeath2) 

# Calcular os erros-padrão robustos a partir da matriz de variância-covariância
robust_se_values2 <- sqrt(diag(robust_se2)) 

# Calcular as razões de prevalência (exponenciação dos coeficientes do modelo)
PR2 <- exp(coef_estimates2)

# Calcular os intervalos de confiança de 95% para as razões de prevalência
lower_CI2 <- exp(coef_estimates2 - 1.96 * robust_se_values2)
upper_CI2 <- exp(coef_estimates2 + 1.96 * robust_se_values2)

# Criar uma tabela final com as variáveis, razões de prevalência e seus intervalos de confiança
results_table <- data.frame(
  Variable = names(coef_estimates2),
  Prevalence_Ratio = round(PR2,2),
  Lower_95CI = round(lower_CI2,2),
  Upper_95CI = round(upper_CI2,2)
)

# Teste de deviance
anova(modDeath2, test = "Chisq")

# Cálculo do pseudo R²
1 - (modDeath2$deviance / modDeath2$null.deviance) #25,36%

