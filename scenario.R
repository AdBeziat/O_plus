## Les instructions sont donnees par les lignes commencant par '##'

## installer les packages necessaires a l'execution du programme (necssaire uniquement la 1ere fois)

install.packages("tidyr")
install.packages("dplyr")
install.packages("openxlsx")
install.packages("stringr")
install.packages("reshape")
install.packages("stats")
install.packages("surveysd")
install.packages("data.table")
install.packages("rlang")
install.packages("mipfp")

## lancer les packages necessaires a l'execution du programme

options(scipen = 999)

library(openxlsx)
library(tidyr)
library(stringr)
library(dplyr)
library(reshape)
library(stats)
library(surveysd)
library(data.table)
library(rlang)
library(mipfp)

## D?finir le chemin d'acces aux fichiers : modifier uniquement la partie entre guillemet

chemin <- "myway" # changer le chemin d'acces au fichier

## Donnez le nom du fichier d'input : modifier uniquement la partie entre guillement
# par exemple, si vous avez nomme votre scenario "scenario1.xlsx", il faut modifier la partie entre guillemet comme suit : "scenario1"

fichier_b2c <- "scenario1"
fichier_b2b <- "scenario1"

## Executer les programmes

#
#
# ATTENTION > Soyez sur.e d'avoir parametre correctement les donnees d'input avant de lancer la commande
#
#

source(paste0(chemin,"/data/programme_b2c.R"))
source(paste0(chemin,"/data/programme_b2b.R"))

