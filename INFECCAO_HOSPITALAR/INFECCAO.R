
# Pacotes usados

require(openxlsx)
require(dplyr)
require(tidyr)

# Definindo o ambiente de trabalho
# Caminho relativo que vai dereto à pasta em que está os arquivos a serem trabalhados

setwd("../INFECCAO_HOSPITALAR/")

# Importando os arquivos

a20 <- loadWorkbook("Dados/indicadores_2020.xlsx")
a21 <- loadWorkbook("Dados/indicadores_2021.xlsx")

#c = list(a2019 = a)
infeccao_banco = list(a2020 = a20, a2021 = a21)
#names(c) = c("A19", "A20")

rm(list = setdiff(ls(), c("infeccao_banco")))

dados = data.frame()
cirurgia = data.frame()
infeccao = data.frame()

# For para ler as tabelas que estão desorganizada

#z = names(infeccao_banco)[3] # [3] apenas para acessar os nomes das planilhas
#sheetNames <- sheets(infeccao_banco[[z]])[-1]
#i = sheetNames[1]

for(z in names(infeccao_banco)){
  
  sheetNames <- sheets(infeccao_banco[[z]])[-1]
  
  for(i in sheetNames)
  {
    
    #$ Critério de eligibilidade Total > 0
    aux1 = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(1,8), rows = 1:12)
    
    if( sum(aux1$Total) > 0)
    {
      
    aux = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(1 : 10), rows = c(1 : 13))
    aux$Mês = i
    aux$Ano = substring(z, 2)
    dados = plyr::rbind.fill(dados, aux)
    
    auxC = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(13 : 16), rows = c(1 : 4))
    auxC$Mês = i
    auxC$Ano = substring(z, 2)
    cirurgia = plyr::rbind.fill(cirurgia, auxC)
    
    auxI = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(13 : 16), rows = c(16 : 24))
    auxI$Mês = i
    auxI$Ano = substring(z, 2)
    infeccao = plyr::rbind.fill(infeccao, auxI)
    
    }
  }
}

rm(i, z, aux, auxC, auxI, aux1)

# Retirar linhas com nome Total da coluna tipo, pois já há uma coluna com os totais

cirurgia = cirurgia %>%
  filter(Tipo != "Total")

# Criando coluna de taxa de infecção cirúrgica

cirurgia = cirurgia %>%
  mutate(
    `Taxa de infecção cirúrgica` = round( (Número.de.infecções / Número.de.cirurgias)*100 , 2)
  )

# Criando o banco de miroorganismos nas unidades

micro = data.frame()

sheetNames <- sheets(infeccao_banco[[1]])[-1]

j = c(1:11)

 #z = names(infeccao_banco)[3]
 #k = sheetNames[1]
 #j = 1


 for(z in names(infeccao_banco)){
   
   for(k in sheetNames){
     
     for(i in j){
       
       # j = which(i == x) # localizacao da clinica no vertor x
       
       aux = readWorkbook(infeccao_banco[[z]], sheet = k, cols = c(1 : 8), rows = c(2 : 13) + 14*i)
       
       aux1 = readWorkbook(infeccao_banco[[z]], sheet = k, cols = c(1,8), rows = 1:12)
       
       if( sum(aux1$Total)>0)
       {
         
         aux1 = unlist(aux1[,1], use.names = F)
         
         aux$Unidade.Funcional = aux1[i]
         aux$Mês = k
         aux$Ano = substring(z, 2)
         
         micro = plyr::rbind.fill(micro, aux)
       }
       
     }
     
   }
   
 }
 
 
micro = micro %>%
  filter(!is.na(MICROORGANISMOS)) %>%
  filter(MICROORGANISMOS != "TOTAL")

#$ Selecionar somente os Microorganismos

micro_indicador = micro %>% 
  select(
    MICROORGANISMOS, Total, Unidade.Funcional, Mês, Ano
  )

#$ __________________________________________________

rm(list = setdiff(ls(), c("dados","micro","cirurgia","infeccao", "micro_indicador")))

#$ __________________________________________________

# Taxa Infecção: Hospital, Por Unidade Funcional, por mês

# Densidade Infecção (Paciente-Dia): Hospital, por Unidade Funcional, por mês

# Distribuição n e % de IRAS: Hospital, por Unidade Funcional, por mês

# Distribuição dos microorganismos isolados


# Organizar Mês e Ano/Mês nos bancos de dados
# Recode

nmes = function(dados, Mês, Ano)
  {
  dados = dados %>% 
  mutate(Mês = 
           case_when
         (
           Mês == "Janeiro" ~ "01",
           Mês == "Fevereiro" ~ "02",
           Mês == "Março" ~ "03",
           Mês == "Abril" ~ "04",
           Mês == "Maio" ~ "05",
           Mês == "Junho" ~ "06",
           Mês == "Julho" ~ "07",
           Mês == "Agosto" ~ "08",
           Mês == "Setembro" ~ "09",
           Mês == "Outubro" ~ "10",
           Mês == "Novembro" ~ "11",
           Mês == "Dezembro" ~ "12"
         )
         )
  dados = dados %>% 
    mutate(
    `Mês/Ano` = as.Date(paste("01", Mês, Ano, sep = "/"), format = "%d/%m/%Y"),
    ANOMES = paste0(format(`Mês/Ano`, "%Y"), format(`Mês/Ano`, "%m"))
    )
  return(dados)
  }

cirurgia = nmes(cirurgia, Mês, Ano)
micro =    nmes(micro, Mês, Ano)
infeccao = nmes(infeccao, Mês, Ano)
dados =    nmes(dados, Mês, Ano)
micro_indicador = nmes(micro_indicador, Mês, Ano)

rm(nmes)

# Criar banco de dados Unidade Funcional

unidade_funcional = as.character(unique(dados$Unidade.Funcional))
unidade_funcional = data.frame(unidade_funcional, stringsAsFactors = F)

unidade_funcional = unidade_funcional %>%
  mutate(
    und_abv = unidade_funcional,
    und_abv = recode(und_abv, 
                     "Unidade A" = "UA", 
                     "Unidade B" = "UB", 
                     "Unidade C" = "UC",
                     "Unidade D" = "UD",
                     "Unidade E" = "UE", 
                     "Unidade F" = "UF", 
                     "Unidade G" = "UG", 
                     "Unidade H" = "UH", 
                     "Unidade I" = "UI", 
                     "Unidade J" = "UJ", 
                     "Unidade K" = "UK",
                     "Unidade L" = "UL")
    
  )

# Criando Banco de dados calendário

calendario = dados %>%
  select(`Mês/Ano`) %>%
  mutate(dupli = duplicated(`Mês/Ano`)) %>%
  filter(dupli == F) %>%
  mutate(dupli = NULL) %>%
  mutate(
    Ano = format(`Mês/Ano`, "%Y"),
    ANOMES = paste0(format(`Mês/Ano`, "%Y"), format(`Mês/Ano`, "%m")))
  
# Retirando colunas desnecessárias

iras = dados %>% 
  select(-c(Total, População, Paciente.Dia))

# Wide to long 

iras = iras %>% 
  gather(Infecção, valor, -c(Unidade.Funcional, Ano, `Mês/Ano`, ANOMES, Mês ))

# Ordenar

iras = 
  iras %>% 
  arrange(Unidade.Funcional, `Mês/Ano`, Infecção)

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Juntando banco de dados iras e cirurgias

auxcirur = cirurgia %>%
  select(-c(Número.de.infecções, Número.de.cirurgias, Total)) 

auxcirur = auxcirur %>%
  rename(Valor = `Taxa de infecção cirúrgica`, Indicador = Tipo)

auxiras = iras %>%
  rename(Indicador = Infecção, Valor = valor)

micro_indicador = micro_indicador %>%
  rename(Indicador = MICROORGANISMOS, Valor = Total)
  
micro_indicador = plyr::rbind.fill(micro_indicador, auxiras)  
micro_indicador = plyr::rbind.fill(micro_indicador, auxcirur)  
rm(auxiras, auxcirur)

# Organizando banco de dados infecção

infeccao = infeccao %>%
  gather(unidade_funcional, valor, -c(Indicador, Mês, Ano, `Mês/Ano`, ANOMES))

infeccao = infeccao %>%
  spread(Indicador, valor)

infeccao = infeccao %>%
  select(Mês:unidade_funcional, IRAS, `PAC/dia`, ICS, `CVC/dia`, PAV, `VM/dia`, ITU, `SVD/dia`)

infeccao = infeccao %>%
  mutate(
    unidade_funcional = 
      recode(unidade_funcional,
             "Unidade.G" = "Unidade G",
             "Unidade.K" = "Unidade K",
             "Unidade.N" = "Unidade N")
    )

# Organizando os banco de dadosa

cirurgia = cirurgia %>% 
  rename(
    `Número de infecções` = Número.de.infecções,
    `Número de cirurgias` = Número.de.cirurgias)

iras = iras %>% 
  rename(`Unidade Funcional` = Unidade.Funcional,
         `Número de infecções` = valor
         #,`Paciente Dia` = Paciente.Dia
         )

micro = micro %>% 
  rename(`Unidade Funcional` = Unidade.Funcional)

unidade_funcional = unidade_funcional %>%
  rename(`Unidade Funcional` = unidade_funcional,
         `.Unidade Funcional` = und_abv)

dados = dados %>% 
  rename(`Número de infecções` = Total, `Paciente Dia` = Paciente.Dia, `Unidade Funcional` = Unidade.Funcional)

infeccao = infeccao %>%
  rename(`Unidade Funcional` = unidade_funcional)

micro_indicador = micro_indicador %>%
  rename(`Unidade Funcional` = Unidade.Funcional)

# Criando banco de dados data de atualização do reatório

datarelatorio = as.character(Sys.time(), format="%d/%m/%Y %H:%M:%S")
datarelatorio = data.frame(datarelatorio, stringsAsFactors = F)

#$ __________________________________________________

save.image("RData/INFECCAO.RData")

# Remover tudo do environment

rm(list = ls())
