
# Pacotes usados

library(dplyr)
library(openxlsx)
require(lubridate)

# Definindo o ambiente de trabalho
# Caminho relativo que vai dereto � pasta em que est� os arquivos a serem trabalhados

setwd("../RESIDENTES/")

RESIDENTES = read.xlsx("Dados/PERFIL_RESIDENTES.xlsx", sheet = 1, startRow = 2)
RESIDENTES_MULTI = read.xlsx("Dados/PERFIL_RESIDENTES.xlsx", sheet = 2, startRow = 2)

#$ Criando coluna de identificador RESIDENTES, RESIDENTES_MULTI

RESIDENTES$`tipo de resid�ncia` = "RESIDENTES"
RESIDENTES_MULTI$`tipo de resid�ncia` = "RESIDENTES_MULTI"

RESIDENTES <- rbind(RESIDENTES, RESIDENTES_MULTI)

rm(RESIDENTES_MULTI)

# Criando vari�vel para contar os residentes

RESIDENTES$`N�mero de residentes` <- 1

# Renomear as colunas 

RESIDENTES = RESIDENTES %>% 
   rename(Programa = PROGRAMA)

RESIDENTES = RESIDENTES %>% 
rename(Hospital = HOSPITAL)

RESIDENTES = RESIDENTES %>% 
  rename(Residentes = RESIDENTES)

RESIDENTES = RESIDENTES %>% 
  rename(Ano = ANO.DE.RESID�NCIA)

RESIDENTES = RESIDENTES %>% 
  rename(Status = STATUS)

# Extraindo amostra

RESIDENTES = RESIDENTES[sample(212), ]

# Salvando

save.image("RData/RESIDENTES.RData")

#load("X:/USID/ESTAGIO_ESTATISTICA/0_Projetos_portif�lio/RESIDENTES/RData/RESIDENTES.RData")