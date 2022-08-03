
# Pacotes usados

library(dplyr)
library(openxlsx)
require(lubridate)

# Definindo o ambiente de trabalho
# Caminho relativo que vai dereto à pasta em que está os arquivos a serem trabalhados

setwd("../RESIDENTES/")

RESIDENTES = read.xlsx("Dados/PERFIL_RESIDENTES.xlsx", sheet = 1, startRow = 2)
RESIDENTES_MULTI = read.xlsx("Dados/PERFIL_RESIDENTES.xlsx", sheet = 2, startRow = 2)

#$ Criando coluna de identificador RESIDENTES, RESIDENTES_MULTI

RESIDENTES$`tipo de residência` = "RESIDENTES"
RESIDENTES_MULTI$`tipo de residência` = "RESIDENTES_MULTI"

RESIDENTES <- rbind(RESIDENTES, RESIDENTES_MULTI)

# Criando variável para contar os residentes

RESIDENTES$`Número de residentes` <- 1

# Renomear as colunas 

RESIDENTES = RESIDENTES %>% 
   rename(Programa = PROGRAMA)

RESIDENTES = RESIDENTES %>% 
rename(Hospital = HOSPITAL)

RESIDENTES = RESIDENTES %>% 
  rename(Residentes = RESIDENTES)

RESIDENTES = RESIDENTES %>% 
  rename(Ano = ANO.DE.RESIDÊNCIA)

RESIDENTES = RESIDENTES %>% 
  rename(Status = STATUS)

# Salvando

save.image("RData/RESIDENTES.RData")

#load("X:/USID/ESTAGIO_ESTATISTICA/0_Projetos_portifólio/RESIDENTES/RData/RESIDENTES.RData")
