#### Importation fichiers bruts population ####

pop <- fread(paste0(chemin, "data/FD_INDCVI_2019.csv"), 
               select = c("NUMMI", "CS1", "AGED", "INPER", "LPRM", "SURF", "TYPL", "NBPI", "STOCD", "ACHLR", "MODV", 
                          "ILETUD", "DEPT", "DIPL", "IRIS", "IPONDI", "CANTVILLE"))

COM <- as.character(read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                 rows = (2:3), cols = (6)))
DEP <- str_sub(COM, 1, 2)

pop <- filter(pop, DEPT == DEP) # filtre dept

popetudiante <- pop # on prepare un second fichier pour filtrer les logements etudiants

popetudiante <- filter(popetudiante, NBPI == "ZZ") # filtre sur les residents hors logements ordinaires
popetudiante <- filter(popetudiante, MODV == 70) # filtre sur les personnes hors menages
popetudiante <- filter(popetudiante, ILETUD != "Z") # filtre sur les etudiants
popetudiante <- filter(popetudiante, DIPL == "13" | DIPL == "14" | DIPL == "15" | DIPL == "16" | DIPL == "17" | 
                         DIPL == "18" | DIPL == "19") # filtre sur les etudiants du secondaire

pop$RES_ETUD <- "NON"

pop <- filter(pop, ACHLR == "6" | ACHLR == "7") # filtre sur l'annee d'achat

pop$NUMMI <- as.numeric(as.character(pop$NUMMI))

popetudiante$NUMMI <- ((max(pop$NUMMI)+1):(nrow(popetudiante)+max(pop$NUMMI)))
popetudiante$RES_ETUD <- "OUI"

pop <- rbind(pop, popetudiante) # on combine pop gen et population etudiante

remove(popetudiante)


# fonction pour eviter les problemes d'arrondi

scale_to_sum_int <- function(x, total, na_to_zero = TRUE) {
  stopifnot(total >= 0)
  if (na_to_zero) x[is.na(x)] <- 0
  if (any(x < 0)) stop("x ne doit pas contenir de valeurs négatives.")
  s <- sum(x)
  if (s == 0) return(integer(length(x)))
  y_cont <- x / s * total
  y <- floor(y_cont)
  k <- total - sum(y)
  if (k > 0) {
    rema <- y_cont - y
    idx <- order(rema, decreasing = TRUE, method = "radix")[seq_len(k)]
    y[idx] <- y[idx] + 1L
  }
  as.integer(y)
}



#### Creation du dataframe menage ####

menages <- filter(pop, LPRM == "1" | RES_ETUD == "OUI")

menages$NUMMI <- paste0(menages$NUMMI, "_", menages$CANTVILLE)

menages$CS1 <- as.character(recode_factor(menages$CS1, "1" = "CE/ART/COM", "2" = "CE/ART/COM", "3" = "CADRE", "4" = "INTERM",
                       "5" = "EMPLOYE", "6" = "OUVRIER", "7" = "RETRAITE", "8" = "INACTIF"))

menages$CL_AGE <- "35-49 ans"
menages$CL_AGE[which(menages$AGED < 35)] <- "18-34 ans"
menages$CL_AGE[which(menages$AGED > 64)] <- "65 ans +"
menages$CL_AGE[which(menages$AGED < 65 & menages$AGED > 49)] <- "50-64 ans"

menages$INPER <- as.character(menages$INPER)
menages$INPER[which(menages$INPER == "Z")] <- "1"
menages$TAILMEN <- as.numeric(as.character(menages$INPER))
menages$TAILMEN[which(menages$TAILMEN > 4)] <- 4
menages$TAILMEN <- as.character(menages$TAILMEN)

menages$POIDS <- as.numeric(as.character(menages$IPONDI))

menages$NBPI <- as.character(recode_factor(menages$NBPI, "01" = "1P", "02" = "2P", "03" = "3P", "04" = "4P", "05" = "5P",
                        "06" = "5P", "07" = "5P", "08" = "5P", "09" = "5P", "10" = "5P",
                        "11" = "5P", "12" = "5P", "13" = "5P", "14" = "5P", "15" = "5P",
                        "16" = "5P", "17" = "5P", "18" = "5P", "19" = "5P", "20" = "5P", "ZZ" = "1P"))


menages$STOCD <- as.character(recode_factor(menages$STOCD, "10" = "Proprietaire", "21" = "Locataire", 
                                            "22" = "Locataire social", 
                                            "23" = "Locataire", "30" = "Locataire"))
menages$STOCD[which(menages$RES_ETUD == "OUI")] <- "Locataire etudiant"

menages$SURF <- as.numeric(as.character(recode_factor(menages$SURF, "1" = "25", "2" = "35", "3" = "50", 
                                                      "4" = "70", "5" = "90",
                                                      "6" = "110", "7" = "135", "Z" = "25")))

menages$CODE_INSEE <- str_sub(menages$IRIS, 1, 5)

menages$COUNT <- "COUNT"

menages_cc <- filter(menages, CODE_INSEE == COM)

CANTON <- as.character(menages_cc[1, "CANTVILLE"])

menages_cant <- filter(menages, CANTVILLE == CANTON)

menages_dep <- menages

remove(pop, menages)


#
#### Taille moyenne logement au departement ####

# calcul de la taille moyenne d'un logement dans le departement au cas ou la commune concerne ne contient pas tous les types de logements

nbpi <- cast(menages_dep, NBPI ~ COUNT, value = "SURF", mean)
nbpi$COUNT <- round(nbpi$COUNT, 1)
nbpi$NBPI <- as.character(nbpi$NBPI)

surf_log_dept_nbpi <- as.data.frame(c("1P", "2P", "3P", "4P", "5P"))
names(surf_log_dept_nbpi) <- "NBPI"
surf_log_dept_nbpi$COUNT <- nbpi$COUNT[match(surf_log_dept_nbpi$NBPI, nbpi$NBPI)]


stocd <- cast(menages_dep, STOCD ~ COUNT, value = "SURF", mean)
stocd$COUNT <- round(stocd$COUNT, 1)
stocd$STOCD <- as.character(stocd$STOCD)

surf_log_dept_stocd <- as.data.frame(c("Locataire", "Locataire social", "Proprietaire", "Locataire etudiant"))
names(surf_log_dept_stocd) <- "STOCD"
surf_log_dept_stocd$COUNT <- stocd$COUNT[match(surf_log_dept_stocd$STOCD, stocd$STOCD)]

remove(stocd, nbpi)


#### chiffre fourni par l'utilisateur : surface et nombre total de logements ####

NB_TOTAL_LOGEMENTS <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                                rows = (7:8), cols = (6))
NB_TOTAL_LOGEMENTS <- NB_TOTAL_LOGEMENTS$Nb
if(is_empty(NB_TOTAL_LOGEMENTS)){NB_TOTAL_LOGEMENTS <- 0}

SURFACE_TOTALE_LOGEMENTS <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                                      rows = (5:6), cols = (6))
SURFACE_TOTALE_LOGEMENTS <- SURFACE_TOTALE_LOGEMENTS$Nb


#### Avec les infos par nombre de pieces #####

# surface moyenne des logements d'apres l'INSEE

menages_cc$COUNT <- "COUNT"

surf_log_com_nbpi <- cast(menages_cc, NBPI ~ COUNT, value = "SURF", mean)
surf_log_com_nbpi$COUNT <- round(surf_log_com_nbpi$COUNT, 1)

if(nrow(surf_log_com_nbpi) == 0){
  
  surf_log_com_nbpi <- surf_log_dept_nbpi$COUNT[match(surf_log_com_nbpi$NBPI, surf_log_dept_nbpi$NBPI)]
  surf_log_com_nbpi$COUNT[is.na(surf_log_com_nbpi$COUNT)] <- surf_log_com_nbpi$COUNT2[is.na(surf_log_com_nbpi$COUNT)]
  surf_log_com_nbpi <- select(surf_log_com_nbpi, -COUNT2)

}

surf_tot_com_nbpi <- cast(menages_cc, NBPI ~ COUNT, value = "SURF", sum)
nb_tot_com_nbpi <- cast(menages_cc, NBPI ~ COUNT, value = "IPONDI", sum)

remove(surf_log_dept_nbpi)

# tableau fourni par lutilisateur - nombre de logements par nombre de pieces

nbpi_nb <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                 rows = (12:17), cols = (5:6))
names(nbpi_nb) <- c("NBPI", "NB")
nbpi_nb$NB[is.na(nbpi_nb$NB)] <- 0
nbpi_nb$NBPI <- as.character(recode_factor(nbpi_nb$NBPI, "Studio" = "1P", 
                                                     "2 pieces" = "2P", "3 pieces" = "3P", 
                                                     "4 pieces" = "4P", "5 pieces et +" = "5P"))

# tableau fourni par lutilisateur - m2 

nbpi_m2 <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                     rows = (20:25), cols = (5:6))
names(nbpi_m2) <- c("NBPI", "SURF")
nbpi_m2$NBPI <- as.character(recode_factor(nbpi_m2$NBPI, "Studio" = "1P", "2 pieces" = "2P", 
                                                     "3 pieces" = "3P",  "4 pieces" = "4P", "5 pieces et +" = "5P"))
nbpi_m2$SURF[is.na(nbpi_m2$SURF)] <- 0
nbpi_m2$NB <- nbpi_m2$SURF / surf_log_com_nbpi$COUNT

# CAS 1 : on a le nombre de m² par nb de pieces, mais pas le nombre de logements, et mais on a le nombre total de logement

if(sum(nbpi_nb$NB) == 0 & sum(nbpi_m2$NB) != 0 & NB_TOTAL_LOGEMENTS != 0){
  
  nbpi_m2$NB <- scale_to_sum_int(nbpi_m2$NB, NB_TOTAL_LOGEMENTS)
  
  
}

# CAS 2 : on a le nombre de m² par nb de pieces, mais pas le nombre de logements, et pas le nombre total de logement

if(sum(nbpi_nb$NB) == 0 & sum(nbpi_m2$NB) != 0 & NB_TOTAL_LOGEMENTS == 0){
  
  nbpi_m2$NB <- round(nbpi_m2$NB, 0)
  
  NB_TOTAL_LOGEMENTS <- sum(nbpi_m2$NB)
  
}

# CAS 3 : on a pas le nombre de m² par nb de pieces, mais on a le nombre de logements

if(sum(nbpi_nb$NB) != 0 & sum(nbpi_m2$NB) == 0){
  
  nbpi_m2$NB <- nbpi_nb$NB
  nbpi_m2$SURF <- nbpi_m2$NB * surf_log_com_nbpi$COUNT
  
  nbpi_m2$SURF <- scale_to_sum_int(nbpi_m2$SURF, SURFACE_TOTALE_LOGEMENTS)
  
}

nbpi <- nbpi_m2

remove(nbpi_m2, nbpi_nb)


#
#### Avec les infos par type de logement ####

# surface moyenne

surf_log_com_stocd <- cast(menages_cc, STOCD ~ COUNT, value = "SURF", mean)
surf_log_com_stocd$COUNT <- round(surf_log_com_stocd$COUNT, 1)
if(nrow(surf_log_com_stocd) < 5){
  
  surf_log_com_stocd$COUNT2 <- surf_log_dept_stocd$COUNT[match(surf_log_com_stocd$STOCD, surf_log_dept_stocd$STOCD)]
  surf_log_com_stocd$COUNT[is.na(surf_log_com_stocd$COUNT)] <- surf_log_com_stocd$COUNT2[is.na(surf_log_com_stocd$COUNT)]
  surf_log_com_stocd <- select(surf_log_com_stocd, -COUNT2)
  
}

surf_tot_com_stocd <- cast(menages_cc, STOCD ~ COUNT, value = "SURF", sum)
nb_tot_com_stocd <- cast(menages_cc, STOCD ~ COUNT, value = "IPONDI", sum)

remove(surf_log_dept_stocd)

# tableau fourni par lutilisateur - nombre de logements

stocd_nb <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                      rows = (28:32), cols = (5:6))
names(stocd_nb) <- c("STOCD", "NB")
stocd_nb$NB[is.na(stocd_nb$NB)] <- 0

stocd_nb$STOCD <- as.character(recode_factor(stocd_nb$STOCD, "Location logement ordinaire" = "Locataire",
                                "Location logement social" = "Locataire social",
                                "Accession a la propriete" = "Proprietaire",
                                "Logement etudiant" = "Locataire etudiant"))

stocd_nb <- arrange(stocd_nb, STOCD)

# tableau de contingence fourni par lutilisateur - m2 

stocd_m2 <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                      rows = (35:39), cols = (5:6))
names(stocd_m2) <- c("STOCD", "SURF")
stocd_m2$SURF[is.na(stocd_m2$SURF)] <- 0
stocd_m2$STOCD <- as.character(recode_factor(stocd_m2$STOCD, "Location logement ordinaire" = "Locataire",
                                "Location logement social" = "Locataire social",
                                "Accession a la propriete" = "Proprietaire",
                                "Logement etudiant" = "Locataire etudiant"))
stocd_m2 <- arrange(stocd_m2, STOCD)

stocd_m2$NB <- stocd_m2$SURF / surf_log_com_stocd$COUNT

# CAS 1 : on a le nombre de m², mais pas le nombre de logements, et mais on a le nombre total de logement

if(sum(stocd_nb$NB) == 0 & sum(stocd_m2$NB) != 0 & NB_TOTAL_LOGEMENTS != 0){
  
  stocd_m2$NB <- scale_to_sum_int(stocd_m2$NB, NB_TOTAL_LOGEMENTS)
  
}

# CAS 2 : on a le nombre de m², mais pas le nombre de logements, et pas le nombre total de logement

if(sum(stocd_nb$NB) == 0 & sum(stocd_m2$NB) != 0 & NB_TOTAL_LOGEMENTS == 0){
  
  stocd_m2$NB <- round(stocd_m2$NB, 0)
  
  NB_TOTAL_LOGEMENTS <- sum(stocd_m2$NB)
  
}

# CAS 3 : on a pas le nombre de m², mais on a le nombre de logements

if(sum(stocd_nb$NB) != 0 & sum(stocd_m2$NB) == 0){
  
  stocd_m2$NB <- stocd_nb$NB
  stocd_m2$SURF <- stocd_m2$NB * surf_log_com_stocd$COUNT
  
  stocd_m2$SURF <- scale_to_sum_int(stocd_m2$SURF, SURFACE_TOTALE_LOGEMENTS)
  
}

stocd <- stocd_m2

remove(stocd_m2, stocd_nb)


#### completer les infos manquantes ####

stocd_nbpi <- cast(menages_cc, STOCD ~ NBPI, value = "IPONDI", sum)
stocd_nbpi$PC_1P <- stocd_nbpi$`1P` / sum(stocd_nbpi$`1P`)
stocd_nbpi$PC_2P <- stocd_nbpi$`2P` / sum(stocd_nbpi$`2P`)
stocd_nbpi$PC_3P <- stocd_nbpi$`3P` / sum(stocd_nbpi$`3P`)
stocd_nbpi$PC_4P <- stocd_nbpi$`4P` / sum(stocd_nbpi$`4P`)
stocd_nbpi$PC_5P <- stocd_nbpi$`5P` / sum(stocd_nbpi$`5P`)

nbpi_stocd <- cast(menages_cc, NBPI ~ STOCD, value = "IPONDI", sum)
nbpi_stocd$PC_Locataire <- nbpi_stocd$Locataire / sum(nbpi_stocd$Locataire)
nbpi_stocd$PC_Locataire_etudiant <- nbpi_stocd$`Locataire etudiant` / sum(nbpi_stocd$`Locataire etudiant`)
nbpi_stocd$PC_Locataire_social <- nbpi_stocd$`Locataire social` / sum(nbpi_stocd$`Locataire social`)
nbpi_stocd$PC_Proprietaire <- nbpi_stocd$Proprietaire / sum(nbpi_stocd$Proprietaire)

# CAS N°1 : on a aucune information pour le nombre de logements par nb de piece, mais on a l'info par type de logement

if(sum(nbpi$NB) == 0 & sum(stocd$NB) != 0 & NB_TOTAL_LOGEMENTS != 0){
  
  nbpi$NB[which(nbpi$NBPI == "1P")] <- (stocd$NB[which(stocd$STOCD == "Locataire")] * stocd_nbpi$PC_1P[which(stocd_nbpi$STOCD == "Locataire")]) + (stocd$NB[which(stocd$STOCD == "Locataire etudiant")] * stocd_nbpi$PC_1P[which(stocd_nbpi$STOCD == "Locataire etudiant")]) + (stocd$NB[which(stocd$STOCD == "Locataire social")] * stocd_nbpi$PC_1P[which(stocd_nbpi$STOCD == "Locataire social")]) + (stocd$NB[which(stocd$STOCD == "Proprietaire")] * stocd_nbpi$PC_1P[which(stocd_nbpi$STOCD == "Proprietaire")])
  nbpi$NB[which(nbpi$NBPI == "2P")] <- (stocd$NB[which(stocd$STOCD == "Locataire")] * stocd_nbpi$PC_2P[which(stocd_nbpi$STOCD == "Locataire")]) + (stocd$NB[which(stocd$STOCD == "Locataire etudiant")] * stocd_nbpi$PC_2P[which(stocd_nbpi$STOCD == "Locataire etudiant")]) + (stocd$NB[which(stocd$STOCD == "Locataire social")] * stocd_nbpi$PC_2P[which(stocd_nbpi$STOCD == "Locataire social")]) + (stocd$NB[which(stocd$STOCD == "Proprietaire")] * stocd_nbpi$PC_2P[which(stocd_nbpi$STOCD == "Proprietaire")])
  nbpi$NB[which(nbpi$NBPI == "3P")] <- (stocd$NB[which(stocd$STOCD == "Locataire")] * stocd_nbpi$PC_3P[which(stocd_nbpi$STOCD == "Locataire")]) + (stocd$NB[which(stocd$STOCD == "Locataire etudiant")] * stocd_nbpi$PC_3P[which(stocd_nbpi$STOCD == "Locataire etudiant")]) + (stocd$NB[which(stocd$STOCD == "Locataire social")] * stocd_nbpi$PC_3P[which(stocd_nbpi$STOCD == "Locataire social")]) + (stocd$NB[which(stocd$STOCD == "Proprietaire")] * stocd_nbpi$PC_3P[which(stocd_nbpi$STOCD == "Proprietaire")])
  nbpi$NB[which(nbpi$NBPI == "4P")] <- (stocd$NB[which(stocd$STOCD == "Locataire")] * stocd_nbpi$PC_4P[which(stocd_nbpi$STOCD == "Locataire")]) + (stocd$NB[which(stocd$STOCD == "Locataire etudiant")] * stocd_nbpi$PC_4P[which(stocd_nbpi$STOCD == "Locataire etudiant")]) + (stocd$NB[which(stocd$STOCD == "Locataire social")] * stocd_nbpi$PC_4P[which(stocd_nbpi$STOCD == "Locataire social")]) + (stocd$NB[which(stocd$STOCD == "Proprietaire")] * stocd_nbpi$PC_4P[which(stocd_nbpi$STOCD == "Proprietaire")])
  nbpi$NB[which(nbpi$NBPI == "5P")] <- (stocd$NB[which(stocd$STOCD == "Locataire")] * stocd_nbpi$PC_5P[which(stocd_nbpi$STOCD == "Locataire")]) + (stocd$NB[which(stocd$STOCD == "Locataire etudiant")] * stocd_nbpi$PC_5P[which(stocd_nbpi$STOCD == "Locataire etudiant")]) + (stocd$NB[which(stocd$STOCD == "Locataire social")] * stocd_nbpi$PC_5P[which(stocd_nbpi$STOCD == "Locataire social")]) + (stocd$NB[which(stocd$STOCD == "Proprietaire")] * stocd_nbpi$PC_5P[which(stocd_nbpi$STOCD == "Proprietaire")])
  
  nbpi$NB <- scale_to_sum_int(nbpi$NB, NB_TOTAL_LOGEMENTS)
  
  nbpi$SURF <- nbpi$NB * surf_log_com_nbpi$COUNT
  
  nbpi$SURF <- scale_to_sum_int(nbpi$SURF, SURFACE_TOTALE_LOGEMENTS)
  
}

# CAS N°2 : on a aucune information pour le nombre de logements par type de logement, mais on a l'info par nb de pieces

if(sum(nbpi$NB) != 0 & sum(stocd$NB) == 0 & NB_TOTAL_LOGEMENTS != 0){
  
  stocd$NB[which(stocd$STOCD == "Locataire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire etudiant")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire social")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Proprietaire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "5P")])
  
  stocd$NB <- scale_to_sum_int(stocd$NB, NB_TOTAL_LOGEMENTS)
  
  stocd$SURF <- stocd$NB * surf_log_com_stocd$COUNT
  
  stocd$SURF <- scale_to_sum_int(stocd$SURF, SURFACE_TOTALE_LOGEMENTS)
  
}

# CAS N°3 : on a aucune information sauf le nombre total de logements

if(sum(nbpi$NB) == 0 & sum(stocd$NB) == 0 & NB_TOTAL_LOGEMENTS != 0){
  
  nb_tot_com_nbpi$NB <- scale_to_sum_int(nb_tot_com_nbpi$COUNT, NB_TOTAL_LOGEMENTS)
  nbpi$NB <- nb_tot_com_nbpi$NB[match(nbpi$NBPI, nb_tot_com_nbpi$NBPI)]
  
  nbpi$SURF <- nbpi$NB * surf_log_com_nbpi$COUNT
  nbpi$SURF <- scale_to_sum_int(nbpi$SURF, SURFACE_TOTALE_LOGEMENTS)
  
  stocd$NB[which(stocd$STOCD == "Locataire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire etudiant")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire social")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Proprietaire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "5P")])
  
  stocd$NB <- scale_to_sum_int(stocd$NB, NB_TOTAL_LOGEMENTS)
  
  stocd$SURF <- stocd$NB * surf_log_com_stocd$COUNT
  
  stocd$SURF <- scale_to_sum_int(stocd$SURF, SURFACE_TOTALE_LOGEMENTS)
  
}

# CAS N°4 : on a aucune information sauf le nombre de m²

if(sum(nbpi$NB) == 0 & sum(stocd$NB) == 0 & NB_TOTAL_LOGEMENTS == 0){
  
  surf_tot_com_nbpi$SURF_PROJET <- scale_to_sum_int(surf_tot_com_nbpi$COUNT, SURFACE_TOTALE_LOGEMENTS)
  nbpi$SURF <- surf_tot_com_nbpi$SURF_PROJET[match(nbpi$NBPI, surf_tot_com_nbpi$NBPI)]
  
  nbpi$NB <- round(nbpi$SURF / surf_log_com_nbpi$COUNT, 0)
  NB_TOTAL_LOGEMENTS <- sum(nbpi$NB)
  
  stocd$NB[which(stocd$STOCD == "Locataire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire etudiant")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_etudiant[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Locataire social")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Locataire_social[which(nbpi_stocd$NBPI == "5P")])
  stocd$NB[which(stocd$STOCD == "Proprietaire")] <- (nbpi$NB[which(nbpi$NBPI == "1P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "1P")]) + (nbpi$NB[which(nbpi$NBPI == "2P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "2P")]) + (nbpi$NB[which(nbpi$NBPI == "3P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "3P")]) + (nbpi$NB[which(nbpi$NBPI == "4P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "4P")]) + (nbpi$NB[which(nbpi$NBPI == "5P")] * nbpi_stocd$PC_Proprietaire[which(nbpi_stocd$NBPI == "5P")])
  
  stocd$NB <- scale_to_sum_int(stocd$NB, NB_TOTAL_LOGEMENTS)
  
  stocd$SURF <- stocd$NB * surf_log_com_stocd$COUNT
  
  stocd$SURF <- scale_to_sum_int(stocd$SURF, SURFACE_TOTALE_LOGEMENTS)

  
}

#### On corrige le fichier menage en fonction des 0 eventuels ####

menages_cc$NBPROJET_NBPI <- nbpi$NB[match(menages_cc$NBPI, nbpi$NBPI)]
menages_cc <- filter(menages_cc, NBPROJET_NBPI != 0)

menages_dep$NBPROJET_NBPI <- nbpi$NB[match(menages_dep$NBPI, nbpi$NBPI)]
menages_dep <- filter(menages_dep, NBPROJET_NBPI != 0)


menages_cc$NBPROJET_STOCD <- stocd$NB[match(menages_cc$STOCD, stocd$STOCD)]
menages_cc <- filter(menages_cc, NBPROJET_STOCD != 0)

menages_dep$NBPROJET_STOCD <- stocd$NB[match(menages_dep$STOCD, stocd$STOCD)]
menages_dep <- filter(menages_dep, NBPROJET_STOCD != 0)

#
#### sociodemo fournis par lutilisateur ####

log_age <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                    rows = (42:46), cols = (5:6))
names(log_age) <- c("AGE", "NB")
log_age$NB[is.na(log_age$NB)] <- 0

if(sum(log_age$NB) != 0){
  
  menages_cc$NB_LOG_AGE <- log_age$NB[match(menages_cc$CL_AGE, log_age$AGE)]
  menages_cc <- filter(menages_cc, NB_LOG_AGE != 0)


} # on supprime les valeurs = a 0 dans le fichier de population si necessaire pour faciliter la convergence ifp

log_csp <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                    rows = (49:56), cols = (5:6))
names(log_csp) <- c("CSP", "NB")
log_csp$NB[is.na(log_csp$NB)] <- 0
log_csp$CSP <- as.character(recode_factor(log_csp$CSP, "Chef d'entrep. / Artisan / Commercant" = "CE/ART/COM",
                               "Cadre" = "CADRE", "Profession intermediaire" = "INTERM", 
                               "Employe" = "EMPLOYE", "Ouvrier" = "OUVRIER", "Retraite" = "RETRAITE",
                               "Sans activite" = "INACTIF"))

if(sum(log_csp$NB) != 0){
  
  menages_cc$NB_LOG_CSP <- log_csp$NB[match(menages_cc$CS1, log_csp$CSP)]
  menages_cc <- filter(menages_cc, NB_LOG_CSP != 0)
  
  
} # on supprime les valeurs = a 0 dans le fichier de population si necessaire pour faciliter la convergence ifp


log_tam <- read.xlsx(paste0(chemin, "input/", fichier_b2c,".xlsx"), sheet = 1, 
                    rows = (59:63), cols = (5:6))
names(log_tam) <- c("TAILMEN", "NB")
log_tam$NB[is.na(log_tam$NB)] <- 0
log_tam$TAILMEN <- recode_factor(log_tam$TAILMEN, "1 personne" = "1", "2 personnes" = "2",
                                "3 personnes" = "3", "4 personnes +" = "4")

if(sum(log_tam$NB) != 0){
  
  menages_cc$NB_LOG_TAM <- log_tam$NB[match(menages_cc$TAILMEN, log_tam$TAILMEN)]
  menages_cc <- filter(menages_cc, NB_LOG_TAM != 0)
  
} # on supprime les valeurs = a 0 dans le fichier de population si necessaire pour faciliter la convergence ifp


#### completer les infos manquantes ####

# tableau de contingence extrapole a la commune pour l'age

if(sum(log_age$NB) == 0){
  
  log_insee_age <- cast(menages_cc, CL_AGE ~ COUNT, value = "IPONDI", sum)
  
  log_age$NB <- log_insee_age$COUNT[match(log_age$AGE, log_insee_age$CL_AGE)]
  
  log_age$NB <- scale_to_sum_int(log_age$NB, NB_TOTAL_LOGEMENTS)
  
  remove(log_insee_age)
  
}

# tableau de contingence extrapole a la commune pour la csp

if(sum(log_csp$NB) == 0){
  
  log_insee_csp <- cast(menages_cc, CS1 ~ COUNT, value = "IPONDI", sum)
  
  log_csp$NB <- log_insee_csp$COUNT[match(log_csp$CSP, log_insee_csp$CS1)]
  
  log_csp$NB <- scale_to_sum_int(log_csp$NB, NB_TOTAL_LOGEMENTS)
  
  remove(log_insee_csp)
  
}

# tableau de contingence extrapole a la commune pour l'age

if(sum(log_tam$NB) == 0){
  
  log_insee_tam <- cast(menages_cc, TAILMEN ~ COUNT, value = "IPONDI", sum)
  
  log_tam$NB <- log_insee_tam$COUNT[match(log_tam$TAILMEN, log_insee_tam$TAILMEN)]
  
  log_tam$NB <- scale_to_sum_int(log_tam$NB, NB_TOTAL_LOGEMENTS)
  
  remove(log_insee_tam)
  
}


#### population synthetique de menages ####

menages_cc$calibWeight <- 0

set.seed(1234)

margins_nbpi <- tidyr::uncount(nbpi, NB)
margins_nbpi <- margins_nbpi[,1]
margins_nbpi <- table(margins_nbpi)

margins_stocd <- tidyr::uncount(stocd, NB)
margins_stocd <- margins_stocd[,1]
margins_stocd <- table(margins_stocd)

margins_csp <- tidyr::uncount(log_csp, NB)
margins_csp <- margins_csp[,1]
margins_csp <- table(margins_csp)

margins_age <- tidyr::uncount(log_age, NB)
margins_age <- margins_age[,1]
margins_age <- table(margins_age)

margins_tam <- tidyr::uncount(log_tam, NB)
margins_tam <- margins_tam[,1]
margins_tam <- table(margins_tam)

margins <- list(margins_nbpi, margins_stocd, margins_age, margins_csp, margins_tam)
tgt_list_dims <- list(1, 2, 3, 4, 5)
variables <- select(menages_cc, NBPI, STOCD, CL_AGE, CS1, TAILMEN)

seed_table <- table(variables)
r_ipfp <- Estimate(seed = seed_table, target.list = tgt_list_dims, 
                   target.data = margins,  method = "ipfp")
poids <- as.data.table(r_ipfp$x.hat)

variables$COUNT <- "COUNT"
variables$NB <- 1
variables <- cast(variables, NBPI + STOCD + CL_AGE + CS1 + TAILMEN ~ COUNT, value = "NB", sum)

poids$COMBI <- paste(poids$NBPI, poids$STOCD, poids$CL_AGE, poids$CS1, poids$TAILMEN, sep = "_")
variables$COMBI <- paste(variables$NBPI, variables$STOCD, variables$CL_AGE, variables$CS1, variables$TAILMEN, sep = "_")
menages_cc$COMBI <- paste(menages_cc$NBPI, menages_cc$STOCD, menages_cc$CL_AGE, menages_cc$CS1, menages_cc$TAILMEN, sep = "_")

poids$NB <- variables$COUNT[match(poids$COMBI, variables$COMBI)]
poids$NB[is.na(poids$NB)] <- 0

poids$W <- poids$N / poids$NB

poids$W[is.nan(poids$W)] <- 0

menages_cc$calibWeight <- poids$W[match(menages_cc$COMBI, poids$COMBI)]

menages_cc <- select(menages_cc, -COMBI)

remove(poids, variables, margins, r_ipfp, tgt_list_dims, margins_nbpi, margins_stocd, seed_table, margins_age, margins_csp, margins_tam,
       nb_tot_com_nbpi, nb_tot_com_stocd, nbpi_stocd, stocd_nbpi, surf_log_com_nbpi, surf_log_com_stocd,
       surf_tot_com_nbpi, surf_tot_com_stocd, menages_cant, menages_dep, nbpi, stocd, log_age, log_csp, log_tam)


#### calcul des LAD ####

menages_cc$ADM_AVERAGE <- 13.5 * menages_cc$calibWeight

adm_csp <- read.csv(file = paste0(chemin, "data/ADM_CSP.csv"), sep = ";")
adm_csp_age <- read.csv(file = paste0(chemin, "data/ADM_CSP_CLAGE.csv"), sep = ";")
adm_csp_tam <- read.csv(file = paste0(chemin, "data/ADM_CSP_TAILMEN.csv"), sep = ";")

menages_cc$ADM_CSP <- 0
menages_cc$ADM_CSP_AGE <- 0
menages_cc$ADM_CSP_TAM <- 0

adm_csp$CS1REF <- as.character(adm_csp$CS1REF)
menages_cc$ADM_CSP <- adm_csp$ADM[match(menages_cc$CS1, adm_csp$CS1REF)]

menages_cc$CSP_AGE <- paste(menages_cc$CS1, menages_cc$CL_AGE, sep = "_")
menages_cc$ADM_CSP_AGE <- adm_csp_age$ADM[match(menages_cc$CSP_AGE, adm_csp_age$COMBI)]

adm_csp_tam$TAM <- recode_factor(adm_csp_tam$TAM, "1PERS" = "1", "2PERS" = "2",
                                 "3PERS" = "3", "4PERS" = "4")
adm_csp_tam$COMBI <- paste0(adm_csp_tam$CS1REF, "_", adm_csp_tam$TAM)

menages_cc$CSP_TAILMEN <- paste(menages_cc$CS1, menages_cc$TAILMEN, sep = "_")
menages_cc$ADM_CSP_TAM <- adm_csp_tam$ADM[match(menages_cc$CSP_TAILMEN, adm_csp_tam$COMBI)]

remove(adm_csp, adm_csp_age, adm_csp_tam)

menages_cc$ADM_CSP <- menages_cc$ADM_CSP * menages_cc$calibWeight
menages_cc$ADM_CSP_AGE <- menages_cc$ADM_CSP_AGE * menages_cc$calibWeight
menages_cc$ADM_CSP_TAM <- menages_cc$ADM_CSP_TAM * menages_cc$calibWeight

NB_TOTAL_ADM <- round(mean(sum(menages_cc$ADM_AVERAGE), sum(menages_cc$ADM_CSP), sum(menages_cc$ADM_CSP_AGE), sum(menages_cc$ADM_CSP_TAM)), 0)

tgt_v1 <- cast(menages_cc, CS1 ~ COUNT, value = "ADM_CSP", sum)
tgt_v1$COUNT <- scale_to_sum_int(tgt_v1$COUNT, NB_TOTAL_ADM)
tgt_v1 <- tidyr::uncount(tgt_v1, COUNT)
tgt_v1 <- table(tgt_v1)

tgt_v1_v2 <- cast(menages_cc, CS1 + CL_AGE ~ COUNT, value = "ADM_CSP_AGE", sum)
tgt_v1_v2$COUNT <- scale_to_sum_int(tgt_v1_v2$COUNT, NB_TOTAL_ADM)
tgt_v1_v2 <- tidyr::uncount(tgt_v1_v2, COUNT)
tgt_v1_v2 <- table(tgt_v1_v2)

tgt_v2_v3 <- cast(menages_cc, CS1 + TAILMEN ~ COUNT, value = "ADM_CSP_AGE", sum)
tgt_v2_v3$COUNT <- scale_to_sum_int(tgt_v2_v3$COUNT, NB_TOTAL_ADM)
tgt_v2_v3 <- tidyr::uncount(tgt_v2_v3, COUNT)
tgt_v2_v3 <- table(tgt_v2_v3)

tgt_list_dims <- list(1, c(1, 2), c(1, 3))
tgt_data <- list(tgt_v1, tgt_v1_v2, tgt_v2_v3)

seed_table <- select(menages_cc, CS1, CL_AGE, TAILMEN)
seed_table <- table(seed_table)

r_ipfp <- Estimate(seed = seed_table, target.list = tgt_list_dims,
                   target.data = tgt_data)

poids <- as.data.table(r_ipfp$x.hat)

poids$COMBI <- paste(poids$CS1, poids$CL_AGE, poids$TAILMEN, sep = "_")

remove(r_ipfp, tgt_data, tgt_list_dims, tgt_v1_v2, tgt_v2_v3, seed_table, tgt_v1)

# Application aux menages

menages_cc$COMBI <- paste(menages_cc$CS1, menages_cc$CL_AGE, menages_cc$TAILMEN, sep = "_")

tcd_log <- cast(menages_cc, COMBI ~ COUNT, value = "calibWeight", sum)
tcd_log$COMBI <- as.character(tcd_log$COMBI)
poids$MENAGES <- tcd_log$COUNT[match(poids$COMBI, tcd_log$COMBI)]
poids$MENAGES[is.na(poids$MENAGES)] <- 0
poids$POIDS_MEN <- poids$N / poids$MENAGES

menages_cc$ADM_POND <- poids$POIDS_MEN[match(menages_cc$COMBI, poids$COMBI)]
menages_cc$ADM_POND <- menages_cc$ADM_POND * menages_cc$calibWeight

tcd_log <- cast(menages_cc, CS1 + CL_AGE + TAILMEN + NBPI + STOCD ~ COUNT, value = "calibWeight", sum)
tcd_adm <- cast(menages_cc, CS1 + CL_AGE + TAILMEN + NBPI + STOCD ~ COUNT, value = "ADM_POND", sum)

tcd_log$COUNT <- scale_to_sum_int(tcd_log$COUNT, NB_TOTAL_LOGEMENTS)
tcd_adm$COUNT <- scale_to_sum_int(tcd_adm$COUNT, NB_TOTAL_ADM)

#### output  ####

OUTPUT <- read.xlsx(paste0(chemin, "data/modele output B2C.xlsx"), sheet = 1, rowNames = FALSE)

OUTPUT$VARIABLES[OUTPUT$VARIABLES == 0] <- NA
OUTPUT$NOMBRE[OUTPUT$NOMBRE == 0] <- NA
OUTPUT$LAD[OUTPUT$LAD == 0] <- NA

tcd_log$NB <- "NB"
tcd_adm$NB <- "NB"

OUTPUT[2, "NOMBRE"] <- NB_TOTAL_LOGEMENTS
OUTPUT[2, "LAD"] <- NB_TOTAL_ADM


nbpi <- as.data.frame(OUTPUT[6:10,]) # nb de piece du logement

nbpi$NBPI <- c("1P", "2P", "3P", "4P", "5P")
  
log_modele <- cast(tcd_log, NBPI ~ NB, value = "COUNT", sum)
adm_modele <- cast(tcd_adm, NBPI ~ NB, value = "COUNT", sum)

nbpi$NOMBRE <- log_modele$NB[match(nbpi$NBPI, log_modele$NBPI)]
nbpi$LAD <- adm_modele$NB[match(nbpi$NBPI, adm_modele$NBPI)]

OUTPUT[6:10, "NOMBRE"] <- nbpi$NOMBRE
OUTPUT[6:10, "LAD"] <- nbpi$LAD

remove(nbpi, log_modele, adm_modele)



stocd <- as.data.frame(OUTPUT[14:17,]) # type de logement

stocd$STOCD <- c("Locataire", "Locataire social", "Proprietaire", "Locataire etudiant")

log_modele <- cast(tcd_log, STOCD ~ NB, value = "COUNT", sum)
adm_modele <- cast(tcd_adm, STOCD ~ NB, value = "COUNT", sum)

stocd$NOMBRE <- log_modele$NB[match(stocd$STOCD, log_modele$STOCD)]
stocd$LAD <- adm_modele$NB[match(stocd$STOCD, adm_modele$STOCD)]

OUTPUT[14:17, "NOMBRE"] <- stocd$NOMBRE
OUTPUT[14:17, "LAD"] <- stocd$LAD

remove(stocd, log_modele, adm_modele)




cl_age <- as.data.frame(OUTPUT[21:24,]) # classes d'age

cl_age$CL_AGE <- c("18-34 ans", "35-49 ans", "50-64 ans", "65 ans +")

log_modele <- cast(tcd_log, CL_AGE ~ NB, value = "COUNT", sum)
adm_modele <- cast(tcd_adm, CL_AGE ~ NB, value = "COUNT", sum)

cl_age$NOMBRE <- log_modele$NB[match(cl_age$CL_AGE, log_modele$CL_AGE)]
cl_age$LAD <- adm_modele$NB[match(cl_age$CL_AGE, adm_modele$CL_AGE)]

OUTPUT[21:24, "NOMBRE"] <- cl_age$NOMBRE
OUTPUT[21:24, "LAD"] <- cl_age$LAD

remove(cl_age, log_modele, adm_modele)



csp <- as.data.frame(OUTPUT[28:34,]) # CSP

csp$CS1 <- c("CE/ART/COM", "CADRE", "INTERM", "EMPLOYE", "OUVRIER", "RETRAITE", "INACTIF")

log_modele <- cast(tcd_log, CS1 ~ NB, value = "COUNT", sum)
adm_modele <- cast(tcd_adm, CS1 ~ NB, value = "COUNT", sum)

csp$NOMBRE <- log_modele$NB[match(csp$CS1, log_modele$CS1)]
csp$LAD <- adm_modele$NB[match(csp$CS1, adm_modele$CS1)]

OUTPUT[28:34, "NOMBRE"] <- csp$NOMBRE
OUTPUT[28:34, "LAD"] <- csp$LAD

remove(csp, log_modele, adm_modele)



tailmen <- as.data.frame(OUTPUT[38:41,]) # taille du menage

tailmen$TAILMEN <- c("1", "2", "3", "4")

log_modele <- cast(tcd_log, TAILMEN ~ NB, value = "COUNT", sum)
adm_modele <- cast(tcd_adm, TAILMEN ~ NB, value = "COUNT", sum)

tailmen$NOMBRE <- log_modele$NB[match(tailmen$TAILMEN, log_modele$TAILMEN)]
tailmen$LAD <- adm_modele$NB[match(tailmen$TAILMEN, adm_modele$TAILMEN)]

OUTPUT[38:41, "NOMBRE"] <- tailmen$NOMBRE
OUTPUT[38:41, "LAD"] <- tailmen$LAD

remove(tailmen, log_modele, adm_modele)


write.xlsx(OUTPUT, file = paste0(chemin, "output/", fichier_b2c,"_output.xlsx"), 
           sheetName = "OUTPUT", colClasses = "numeric", append = TRUE)

remove(menages_cc, OUTPUT, poids, tcd_adm, tcd_log, CANTON, COM, DEP, NB_TOTAL_ADM, 
       NB_TOTAL_LOGEMENTS, SURFACE_TOTALE_LOGEMENTS, scale_to_sum_int)









