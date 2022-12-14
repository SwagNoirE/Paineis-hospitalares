
# Pacotes usados

require(openxlsx)
require(dplyr)
require(tidyr)

# Definindo o ambiente de trabalho
# Caminho relativo que vai dereto ? pasta em que est? os arquivos a serem trabalhados

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

# For para ler as tabelas que est?o desorganizada

#z = names(infeccao_banco)[3] # [3] apenas para acessar os nomes das planilhas
#sheetNames <- sheets(infeccao_banco[[z]])[-1]
#i = sheetNames[1]

for(z in names(infeccao_banco)){
  
  sheetNames <- sheets(infeccao_banco[[z]])[-1]
  
  for(i in sheetNames)
  {
    
    #$ Crit?rio de eligibilidade Total > 0
    aux1 = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(1,8), rows = 1:12)
    
    if( sum(aux1$Total) > 0)
    {
      
    aux = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(1 : 10), rows = c(1 : 13))
    aux$M?s = i
    aux$Ano = substring(z, 2)
    dados = plyr::rbind.fill(dados, aux)
    
    auxC = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(13 : 16), rows = c(1 : 4))
    auxC$M?s = i
    auxC$Ano = substring(z, 2)
    cirurgia = plyr::rbind.fill(cirurgia, auxC)
    
    auxI = readWorkbook(infeccao_banco[[z]], sheet = i, cols = c(13 : 16), rows = c(16 : 24))
    auxI$M?s = i
    auxI$Ano = substring(z, 2)
    infeccao = plyr::rbind.fill(infeccao, auxI)
    
    }
  }
}

rm(i, z, aux, auxC, auxI, aux1)

# Retirar linhas com nome Total da coluna tipo, pois j? h? uma coluna com os totais

cirurgia = cirurgia %>%
  filter(Tipo != "Total")

# Criando coluna de taxa de infec??o cir?rgica

cirurgia = cirurgia %>%
  mutate(
    `Taxa de infec??o cir?rgica` = round( (N?mero.de.infec??es / N?mero.de.cirurgias)*100 , 2)
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
         aux$M?s = k
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
    MICROORGANISMOS, Total, Unidade.Funcional, M?s, Ano
  )

#$ __________________________________________________

rm(list = setdiff(ls(), c("dados","micro","cirurgia","infeccao", "micro_indicador")))

#$ __________________________________________________

# Taxa Infec??o: Hospital, Por Unidade Funcional, por m?s

# Densidade Infec??o (Paciente-Dia): Hospital, por Unidade Funcional, por m?s

# Distribui??o n e % de IRAS: Hospital, por Unidade Funcional, por m?s

# Distribui??o dos microorganismos isolados


# Organizar M?s e Ano/M?s nos bancos de dados
# Recode

nmes = function(dados, M?s, Ano)
  {
  dados = dados %>% 
  mutate(M?s = 
           case_when
         (
           M?s == "Janeiro" ~ "01",
           M?s == "Fevereiro" ~ "02",
           M?s == "Mar?o" ~ "03",
           M?s == "Abril" ~ "04",
           M?s == "Maio" ~ "05",
           M?s == "Junho" ~ "06",
           M?s == "Julho" ~ "07",
           M?s == "Agosto" ~ "08",
           M?s == "Setembro" ~ "09",
           M?s == "Outubro" ~ "10",
           M?s == "Novembro" ~ "11",
           M?s == "Dezembro" ~ "12"
         )
         )
  dados = dados %>% 
    mutate(
    `M?s/Ano` = as.Date(paste("01", M?s, Ano, sep = "/"), format = "%d/%m/%Y"),
    ANOMES = paste0(format(`M?s/Ano`, "%Y"), format(`M?s/Ano`, "%m"))
    )
  return(dados)
  }

cirurgia = nmes(cirurgia, M?s, Ano)
micro =    nmes(micro, M?s, Ano)
infeccao = nmes(infeccao, M?s, Ano)
dados =    nmes(dados, M?s, Ano)
micro_indicador = nmes(micro_indicador, M?s, Ano)

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

# Criando Banco de dados calend?rio

calendario = dados %>%
  select(`M?s/Ano`) %>%
  mutate(dupli = duplicated(`M?s/Ano`)) %>%
  filter(dupli == F) %>%
  mutate(dupli = NULL) %>%
  mutate(
    Ano = format(`M?s/Ano`, "%Y"),
    ANOMES = paste0(format(`M?s/Ano`, "%Y"), format(`M?s/Ano`, "%m")))
  
# Retirando colunas desnecess?rias

iras = dados %>% 
  select(-c(Total, Popula??o, Paciente.Dia))

# Wide to long 

iras = iras %>% 
  gather(Infec??o, valor, -c(Unidade.Funcional, Ano, `M?s/Ano`, ANOMES, M?s ))

# Ordenar

iras = 
  iras %>% 
  arrange(Unidade.Funcional, `M?s/Ano`, Infec??o)

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Juntando banco de dados iras e cirurgias

auxcirur = cirurgia %>%
  select(-c(N?mero.de.infec??es, N?mero.de.cirurgias, Total)) 

auxcirur = auxcirur %>%
  rename(Valor = `Taxa de infec??o cir?rgica`, Indicador = Tipo)

auxiras = iras %>%
  rename(Indicador = Infec??o, Valor = valor)

micro_indicador = micro_indicador %>%
  rename(Indicador = MICROORGANISMOS, Valor = Total)
  
micro_indicador = plyr::rbind.fill(micro_indicador, auxiras)  
micro_indicador = plyr::rbind.fill(micro_indicador, auxcirur)  
rm(auxiras, auxcirur)

# Organizando banco de dados infec??o

infeccao = infeccao %>%
  gather(unidade_funcional, valor, -c(Indicador, M?s, Ano, `M?s/Ano`, ANOMES))

infeccao = infeccao %>%
  spread(Indicador, valor)

infeccao = infeccao %>%
  select(M?s:unidade_funcional, IRAS, `PAC/dia`, ICS, `CVC/dia`, PAV, `VM/dia`, ITU, `SVD/dia`)

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
    `N?mero de infec??es` = N?mero.de.infec??es,
    `N?mero de cirurgias` = N?mero.de.cirurgias)

iras = iras %>% 
  rename(`Unidade Funcional` = Unidade.Funcional,
         `N?mero de infec??es` = valor
         #,`Paciente Dia` = Paciente.Dia
         )

micro = micro %>% 
  rename(`Unidade Funcional` = Unidade.Funcional)

unidade_funcional = unidade_funcional %>%
  rename(`Unidade Funcional` = unidade_funcional,
         `.Unidade Funcional` = und_abv)

dados = dados %>% 
  rename(`N?mero de infec??es` = Total, `Paciente Dia` = Paciente.Dia, `Unidade Funcional` = Unidade.Funcional)

infeccao = infeccao %>%
  rename(`Unidade Funcional` = unidade_funcional)

micro_indicador = micro_indicador %>%
  rename(`Unidade Funcional` = Unidade.Funcional)

# Criando banco de dados data de atualiza??o do reat?rio

datarelatorio = as.character(Sys.time(), format="%d/%m/%Y %H:%M:%S")
datarelatorio = data.frame(datarelatorio, stringsAsFactors = F)

#$ __________________________________________________

save.image("RData/INFECCAO.RData")

# Remover tudo do environment

rm(list = ls())
