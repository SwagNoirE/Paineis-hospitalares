
# Pacotes usados

library(dplyr)
library(openxlsx)
require(lubridate)

# Definindo o ambiente de trabalho
# Caminho relativo que vai dereto à pasta em que está os arquivos a serem trabalhados

setwd("../COMISSAO/")

# Lendo xlsx

Comissao = read.xlsx("Comissões.xlsx", sheet = 1, detectDates = T)
composicao = read.xlsx("Comissões.xlsx", sheet = 2)

# Verificar se os nomes das comissões estão na base composicao

table(Comissao$Nome %in% composicao$`Nome.-.Comissão`)
table(composicao$`Nome.-.Comissão` %in% Comissao$Nome)

# Renomeando as variáveis

composicao = 
  composicao %>%
  rename(
    `Nome da comissão` = "Nome.-.Comissão",
    `Membros` = "Nome.-.Membros",
    `Tipo de Membro` = "Tipo.-.Membro")

Comissao = 
  Comissao %>%
  rename(
    `Nome da comissão` = Nome,
    `Data Portaria` = `Data.-.Portaria`,
    `Data Boletim de Serviço` = `Data.-.Boletim.de.Serviço`,
    `Regimento interno` = `Regimento.Interno.-.Boletim.de.Serviço`
    )

names(Comissao) = gsub("\\.", " ", names(Comissao))

#$ Ordenando a base composicao

composicao = 
  composicao %>%
  arrange(`Nome da comissão`, Membros)

#$ Identificando número de pessoas

composicao = 
  composicao %>%
  mutate(
    `Número de colaboradores em comissões` = if_else(duplicated(Membros), 0, 1)
  )

# Criando ID para os membros

composicao = 
  composicao %>%
  mutate(`Código sequencial dos membros` = as.numeric(factor(Membros)))

composicao = 
  composicao %>%
  group_by(Membros) %>%
  mutate(`Número de comissões por membro` = row_number()) %>%
  ungroup()

composicao = data.table::setDT(composicao)[!is.na(Membros), "Número de membros por comissão" := data.table::rleid(Membros), "Nome da comissão"]

# Criando Flag para contar comissões e colaboradores

composicao = 
  composicao %>%
  mutate(
    `Número de membros` = 1
  )

Comissao = 
  Comissao %>%
  mutate(
    `Número de comissões` = 1
  )

# Duracao da data de criação da portaria até agora

Comissao = 
  Comissao %>%
  mutate(
    `Data Portaria` = as.Date(`Data Portaria`, format = "%Y-%m-%d"),
    `Data Hoje` = Sys.Date()) %>%
  mutate(
    `Duração em Mêses Decimal` = as.numeric(difftime(`Data Hoje`, `Data Portaria`, units = "days")/30)
  )

Comissao = 
  Comissao %>%
  mutate(
    `Duração em Mêses` = (interval(`Data Portaria`, `Data Hoje`)) %/% months(1)
    )

# Alterar tipo do campo

Comissao = 
  Comissao %>%
  mutate(`Boletim de Serviço` = as.character(`Boletim de Serviço`))

#$ Reclassificar Regimento interno

Comissao = Comissao %>%
  mutate(`Regimento interno` = 
           case_when(
             `Regimento interno` == "Sim"  ~ "Sim",
             TRUE ~ "Não")
         )

# Extraindo amostra


Comissao = Comissao[sample(81), ]


composicao = composicao[sample(708), ]

# Retirar os valores NA da coluna nome da comissão

Comissao =
  Comissao %>%
  filter(!is.na(`Nome da comissão`))

# Salvando em RData

save.image("Comissoes.RData")

#load("X:/USID/ESTAGIO_ESTATISTICA/0_Projetos_portifólio/COMISSAO/Comissoes.RData")
