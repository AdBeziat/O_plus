#### Importation fichiers etab ####

sirene <- read.csv(paste0(chemin,"data/sirene.csv"), sep = ",")

COM <- as.character(read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                 rows = (2:3), cols = (7))) # on recupere la commune de reference
DPT <- str_sub(COM, 1, 2) # on recupere le departement de la commune de reference

sirene <- dplyr::select(sirene, id:ZONE, EFET_ESTIM:SURFACE)

sirene$COUNT <- "COUNT"

sirene$SURFACE[which(sirene$SURFACE < 1001)] <- round(sirene$SURFACE[which(sirene$SURFACE < 1001)], -1) # on arrondit la surface des etablissements a la dizaine pres (pour les etablissements de < 1000 m²)
sirene$SURFACE[which(sirene$SURFACE > 1000)] <- round(sirene$SURFACE[which(sirene$SURFACE > 1000)], -2) # on arrondit la surface des etablissements a la centaine pres (pour les etablissements de > 1000 m²)
sirene$SURFACE[which(sirene$SURFACE == 0)] <- 10

sirene_cc <- filter(sirene, cc_ad == COM) # les etablissements de la commune de reference
sirene_dep <- filter(sirene, DEP == DPT) # les etablissements du departement de reference

remove(sirene)

#### M2 par activites detailles ####

# on importe les m² d'activites par grandes categories d'activites renseignes par l'utilisateur

activ <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                 rows = (5:8), cols = (7))
names(activ) <- "M2"
activ$ACTIV <- c("A_COM", "B_EQP", "C_BUR")

# on importe les m² d'activites par categorie d'activite detaillee renseignes par l'utilisateur

# pour les commerces

a_com <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                   rows = (13:19), cols = c(5,7))
names(a_com) <- c("ACTIV_DET", "M2")
a_com$M2[is.na(a_com$M2)] <- 0
a_com$ACTIV_DET <- as.character(c("A_PSS", "B_GSS", "C_HRC", "D_GRO", "E_ART", "F_SRP"))

# pour les equipements

b_eqp <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                       rows = (22:27), cols = c(5, 7))
names(b_eqp) <- c("ACTIV_DET", "M2")
b_eqp$M2[is.na(b_eqp$M2)] <- 0
b_eqp$ACTIV_DET <- as.character(c("G_FOP", "H_ENS", "I_SPO", "J_HOP", "K_CUL"))

# pour les bureaux tertiaires

c_bur <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
               rows = (30:33), cols = c(5, 7))
names(c_bur) <- c("ACTIV_DET", "M2")
c_bur$M2[is.na(c_bur$M2)] <- 0
c_bur$ACTIV_DET <- as.character(c("L_TER", "M_BNT", "N_STP"))

# si les m² ne sont pas renseignes pour les types d'activite detailles
# on les extrapole a partir du fichier SIRENE (tableau de frequence des m² par types d'activite detailles)

if(sum(a_com$M2) == 0 & activ$M2[which(activ$ACTIV == "A_COM")] != 0){ # si il y a des commerces mais que l'utilisateur n'a pas fourni le nb de m2 detaille
  
  etab_sirene_correspondant <- filter(sirene_cc, ACTIV == "A_COM")
  etab_sirene_correspondant$COUNT <- "COUNT"
  
  etab_sirene_correspondant <- cast(etab_sirene_correspondant, ACTIV_DET ~ COUNT, value = "SURFACE", sum)
  etab_sirene_correspondant$COUNT <- etab_sirene_correspondant$COUNT / sum(etab_sirene_correspondant$COUNT) # m² des commerces detailles dans le fichier sirene
  
  M2 <- activ$M2[which(activ$ACTIV == "A_COM")] # m² des commerces declares par l'utilisateur
  
  etab_sirene_correspondant$COUNT <- round(etab_sirene_correspondant$COUNT * M2, -1)
  
  a_com$M2 <- etab_sirene_correspondant$COUNT[match(a_com$ACTIV_DET, etab_sirene_correspondant$ACTIV_DET)]
  a_com$M2[is.na(a_com$M2)] <- 0
  
  remove(etab_sirene_correspondant, M2)
  
}

if(sum(b_eqp$M2) == 0 & activ$M2[which(activ$ACTIV == "B_EQP")] != 0){ # si il y a des equipements mais que l'utilisateur n'a pas fourni le nb de m2 detaille
  
  etab_sirene_correspondant <- filter(sirene_cc, ACTIV == "B_EQP")
  etab_sirene_correspondant$COUNT <- "COUNT"
  
  etab_sirene_correspondant <- cast(etab_sirene_correspondant, ACTIV_DET ~ COUNT, value = "SURFACE", sum)
  
  etab_sirene_correspondant$COUNT <- etab_sirene_correspondant$COUNT / sum(etab_sirene_correspondant$COUNT)
  
  M2 <- activ$M2[which(activ$ACTIV == "B_EQP")]
  
  etab_sirene_correspondant$COUNT <- round(etab_sirene_correspondant$COUNT * M2, -1)
  
  b_eqp$M2 <- etab_sirene_correspondant$COUNT[match(b_eqp$ACTIV_DET, etab_sirene_correspondant$ACTIV_DET)]
  
  b_eqp$M2[is.na(b_eqp$M2)] <- 0
  
  remove(etab_sirene_correspondant, M2)
  
}

if(sum(c_bur$M2) == 0 & activ$M2[which(activ$ACTIV == "C_BUR")] != 0){ # si il y a des bureaux mais que l'utilisateur n'a pas fourni le nb de m2 detaille
  
  etab_sirene_correspondant <- filter(sirene_cc, ACTIV == "C_BUR")
  etab_sirene_correspondant$COUNT <- "COUNT"
  
  etab_sirene_correspondant <- cast(etab_sirene_correspondant, ACTIV_DET ~ COUNT, value = "SURFACE", sum)
  
  etab_sirene_correspondant$COUNT <- etab_sirene_correspondant$COUNT / sum(etab_sirene_correspondant$COUNT)
  
  M2 <- activ$M2[which(activ$ACTIV == "C_BUR")]
  
  etab_sirene_correspondant$COUNT <- round(etab_sirene_correspondant$COUNT * M2, -1)
  
  c_bur$M2 <- etab_sirene_correspondant$COUNT[match(c_bur$ACTIV_DET, etab_sirene_correspondant$ACTIV_DET)]
  
  c_bur$M2[is.na(c_bur$M2)] <- 0
  
  remove(etab_sirene_correspondant, M2)
  
}

#

#### Nombre d'etab par activites detailles ####

remove(activ)

# on importe le nombre d'etablissements et d'emplois renseignes par l'utilisateur
# et on concatene avec le df deja cree contenant les m²

# pour les commerces

nb_a_com <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                   rows = (13:19), cols = c(5,8,9))
names(nb_a_com) <- c("ACTIV_DET", "NB", "EMPLOIS")
nb_a_com$NB[is.na(nb_a_com$NB)] <- 0
nb_a_com$EMPLOIS[is.na(nb_a_com$EMPLOIS)] <- 0
nb_a_com$ACTIV_DET <- as.character(c("A_PSS", "B_GSS", "C_HRC", "D_GRO", "E_ART", "F_SRP"))
a_com$NB <- nb_a_com$NB
a_com$EMPLOIS <- nb_a_com$EMPLOIS
remove(nb_a_com)

# pour les equipements

nb_b_eqp <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                   rows = (22:27), cols = c(5,8,9))
names(nb_b_eqp) <- c("ACTIV_DET", "NB", "EMPLOIS")
nb_b_eqp$NB[is.na(nb_b_eqp$NB)] <- 0
nb_b_eqp$EMPLOIS[is.na(nb_b_eqp$EMPLOIS)] <- 0
nb_b_eqp$ACTIV_DET <- as.character(c("G_FOP", "H_ENS", "I_SPO", "J_HOP", "K_CUL"))
b_eqp$NB <- nb_b_eqp$NB
b_eqp$EMPLOIS <- nb_b_eqp$EMPLOIS
remove(nb_b_eqp)

# pour les bureaux

nb_c_bur <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                   rows = (30:33), cols = c(5,8,9))
names(nb_c_bur) <- c("ACTIV_DET", "NB", "EMPLOIS")
nb_c_bur$NB[is.na(nb_c_bur$NB)] <- 0
nb_c_bur$EMPLOIS[is.na(nb_c_bur$EMPLOIS)] <- 0
nb_c_bur$ACTIV_DET <- as.character(c("L_TER", "M_BNT", "N_STP"))
c_bur$NB <- nb_c_bur$NB
c_bur$EMPLOIS <- nb_c_bur$EMPLOIS
remove(nb_c_bur)

# creation du df activ_det = toutes les activites detaillees

activ_det <- rbind(a_com, b_eqp, c_bur)

remove(a_com, b_eqp, c_bur)



#### Estimation du nombre d'etablissement et d'emplois ####

# pour le modele de generation des mouvements, on a besoin d'un nombre d'etablissements et d'un nombre d'emploi
# pour le moment les categories utilisees pour les fonctions de generation sont plus detaillees que les categories renseignees par l'utilisateur d'O+
# on cree un fichier d'etablissement synthetique avec une ponderation fonction des types d'etablissements presents dans le fichier sirene

activ_utilisateurs <- filter(activ_det, M2 > 0 | NB > 0 | EMPLOIS > 0)

x <- 1

while(x <= nrow(activ_utilisateurs)){
  
  TYPE <- activ_utilisateurs[x, "ACTIV_DET"]
  type <- filter(sirene_cc, ACTIV_DET == TYPE)
  if(nrow(type) == 0){type <- filter(sirene_dep, ACTIV_DET == TYPE)}
  
  if(activ_utilisateurs[x, "NB"] == 0){
  
    M2_MOY_TYPE <- mean(type$SURFACE)
    NB_MOY_TYPE <- round(activ_utilisateurs[x, "M2"] / M2_MOY_TYPE, 1)
    activ_det$NB[which(activ_det$ACTIV_DET == TYPE)] <- NB_MOY_TYPE
    
  }else{
    
    M2_MOY_TYPE <- mean(type$SURFACE)
    NB_MOY_TYPE <- activ_utilisateurs[x, "NB"]
    
  }
  
  m2_par_activdet <- cast(type, ST45 ~ COUNT, value = "SURFACE", sum)
  m2_par_activdet$PC <- round(m2_par_activdet$COUNT / sum(m2_par_activdet$COUNT), 3)
  
  emplois_par_activdet <- cast(type, ST45 ~ COUNT, value = "EFET_ESTIM", sum)
  emplois_par_activdet$COUNT <- round(emplois_par_activdet$COUNT, 1)
  emplois_par_activdet$M2 <- m2_par_activdet$COUNT
  emplois_par_activdet$EMPLOIS_PAR_M2 <- emplois_par_activdet$COUNT / emplois_par_activdet$M2
  emplois_par_activdet$M2 <- m2_par_activdet$PC * activ_utilisateurs[x, "M2"]
  emplois_par_activdet$EMPLOIS <- emplois_par_activdet$EMPLOIS_PAR_M2 * emplois_par_activdet$M2
  
  EMP_TOTAL_ESTIM <- activ_utilisateurs[x, "EMPLOIS"]
  
  etab <- dplyr::select(m2_par_activdet, ST45, PC)
  etab$ACTIV_DET <- TYPE
  etab$POND <- NB_MOY_TYPE * etab$PC
  etab$EMPLOI <- emplois_par_activdet$EMPLOIS
  
  if(EMP_TOTAL_ESTIM > 0){
    
    etab$EMP_TOT <- etab$EMPLOI / sum(etab$EMPLOI) * EMP_TOTAL_ESTIM
    etab$EMPLOI <- round(etab$EMP_TOT, 1)
    etab <- dplyr::select(etab, -EMP_TOT)
    
  }else{
    
    activ_det$EMPLOIS[which(activ_det$ACTIV_DET == TYPE)] <- round(sum(etab$EMPLOI), 1)
    
  }
  
  if(!exists("etab_synth")){
    
    etab_synth <- etab
    
  }else{
    
    etab_synth <- rbind(etab, etab_synth)
    
  }
  
  remove(etab, emplois_par_activdet, m2_par_activdet, TYPE, type, M2_MOY_TYPE, NB_MOY_TYPE, EMP_TOTAL_ESTIM)
  
  x <- x + 1
  
}

remove(x, activ_utilisateurs)

etab_synth$EMPLOI_PAR_ETAB <- etab_synth$EMPLOI / etab_synth$POND

#
#### modele de generation des mouvements, fonctions de Beziat, 2017 ####

# generation des mvts

gen <- read.csv(paste0(chemin,"data/TMV_st45.csv"), sep = ",")

etab_synth$MVT <- 0

etab_synth$X <- gen$x[match(etab_synth$ST45, gen$ST45)]
etab_synth$Y <- gen$y[match(etab_synth$ST45, gen$ST45)]
etab_synth$TYPE <- gen$type[match(etab_synth$ST45, gen$ST45)]

etab_synth$EMPLOI_PAR_ETAB[which(etab_synth$EMPLOI_PAR_ETAB < 1.1)] <- 1.1
etab_synth$MVT[which(etab_synth$TYPE == "lin")] <- (etab_synth$EMPLOI_PAR_ETAB[which(etab_synth$TYPE == "lin")] * etab_synth$X[which(etab_synth$TYPE == "lin")]) + etab_synth$Y[which(etab_synth$TYPE == "lin")]
etab_synth$MVT[which(etab_synth$TYPE == "log")] <- (log(etab_synth$EMPLOI_PAR_ETAB[which(etab_synth$TYPE == "log")]) * etab_synth$X[which(etab_synth$TYPE == "log")]) + etab_synth$Y[which(etab_synth$TYPE == "log")]

remove(gen)

etab_synth$MVT_ANNUEL_POND <- etab_synth$MVT * 48 * etab_synth$POND

# livraison

reception_enlevement <- read.csv(paste0(chemin,"data/TYPE_OPE.csv"), sep = ",")

etab_synth$PC_RECEPTION <- reception_enlevement$PC_R[match(etab_synth$ST45, reception_enlevement$ST45)]
etab_synth$REC_ANNUEL_POND <- etab_synth$MVT_ANNUEL_POND * etab_synth$PC_RECEPTION

remove(reception_enlevement)

# poids et conditionnement

type_cargo <- read.csv(paste0(chemin,"data/FRET.csv"), sep = ",")

etab_synth$PC_CARGO <- type_cargo$T1[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT1 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T2[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT2 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T3[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT3 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T4[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT4 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T5[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT5 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T6[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT6 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T7[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT7 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

etab_synth$PC_CARGO <- type_cargo$T8[match(etab_synth$ST45, type_cargo$ST45)]
etab_synth$MVT8 <- etab_synth$PC_CARGO * etab_synth$MVT_ANNUEL_POND

remove(type_cargo)

#### wrap up ####

OUTPUT <- read.xlsx(paste0(chemin, "data/modele output B2B.xlsx"), sheet = 1, rowNames = FALSE)


nb_a_com <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                      rows = (13:19), cols = c(5,8))
names(nb_a_com) <- c("ACTIV_DET", "NB")
nb_a_com$NB <- as.character(c("A_PSS", "B_GSS", "C_HRC", "D_GRO", "E_ART", "F_SRP"))

nb_b_eqp <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                      rows = (22:27), cols = c(5, 8))
names(nb_b_eqp) <- c("ACTIV_DET", "NB")
nb_b_eqp$NB <- as.character(c("G_FOP", "H_ENS", "I_SPO", "J_HOP", "K_CUL"))

nb_c_bur <- read.xlsx(paste0(chemin, "input/", fichier_b2b,".xlsx"), sheet = 1, 
                      rows = (30:33), cols = c(5, 8))
names(nb_c_bur) <- c("ACTIV_DET", "NB")
nb_c_bur$NB <- as.character(c("L_TER", "M_BNT", "N_STP"))

output_det <- rbind(nb_a_com, nb_b_eqp, nb_c_bur)

remove(nb_a_com, nb_b_eqp, nb_c_bur)

etab_synth$COUNT <- "COUNT"
etab_synth$nb <- 1

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "POND", sum)
tcd$COUNT <- ceiling(tcd$COUNT)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$NOMBRE <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$NOMBRE[is.na(OUTPUT$NOMBRE)] <- 0
OUTPUT[15, "NOMBRE"] <- sum(OUTPUT[1:14,"NOMBRE"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "EMPLOI", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- ceiling(tcd$COUNT)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$EMPLOIS <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$EMPLOIS[is.na(OUTPUT$EMPLOIS)] <- 0
OUTPUT[15, "EMPLOIS"] <- sum(OUTPUT[1:14,"EMPLOIS"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT_ANNUEL_POND", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_AN <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_AN[is.na(OUTPUT$LIV_AN)] <- 0
OUTPUT[15, "LIV_AN"] <- sum(OUTPUT[1:14,"LIV_AN"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT1", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_MOINS_1KG <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_MOINS_1KG[is.na(OUTPUT$LIV_MOINS_1KG)] <- 0
OUTPUT[15, "LIV_MOINS_1KG"] <- sum(OUTPUT[1:14,"LIV_MOINS_1KG"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT2", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_3KG <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_3KG[is.na(OUTPUT$LIV_COLIS_MOINS_3KG)] <- 0
OUTPUT[15, "LIV_COLIS_MOINS_3KG"] <- sum(OUTPUT[1:14,"LIV_COLIS_MOINS_3KG"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT3", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_30KG <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_30KG[is.na(OUTPUT$LIV_COLIS_MOINS_30KG)] <- 0
OUTPUT[15, "LIV_COLIS_MOINS_30KG"] <- sum(OUTPUT[1:14,"LIV_COLIS_MOINS_30KG"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT4", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_100KG <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_COLIS_MOINS_100KG[is.na(OUTPUT$LIV_COLIS_MOINS_100KG)] <- 0
OUTPUT[15, "LIV_COLIS_MOINS_100KG"] <- sum(OUTPUT[1:14,"LIV_COLIS_MOINS_100KG"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT5", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_COLIS_PLUS_100KG <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_COLIS_PLUS_100KG[is.na(OUTPUT$LIV_COLIS_PLUS_100KG)] <- 0
OUTPUT[15, "LIV_COLIS_PLUS_100KG"] <- sum(OUTPUT[1:14,"LIV_COLIS_PLUS_100KG"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT6", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_PALETTES <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_PALETTES[is.na(OUTPUT$LIV_PALETTES)] <- 0
OUTPUT[15, "LIV_PALETTES"] <- sum(OUTPUT[1:14,"LIV_PALETTES"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT7", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_AUTRE <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_AUTRE[is.na(OUTPUT$LIV_AUTRE)] <- 0
OUTPUT[15, "LIV_AUTRE"] <- sum(OUTPUT[1:14,"LIV_AUTRE"])

tcd <- cast(etab_synth, ACTIV_DET ~ COUNT, value = "MVT8", sum)
tcd$ACTIV_DET <- as.character(tcd$ACTIV_DET)
tcd$COUNT <- round(tcd$COUNT, 1)
output_det$NOMBRE <- tcd$COUNT[match(output_det$NB, tcd$ACTIV_DET)]
OUTPUT$LIV_LOTS_COMPLETS <- output_det$NOMBRE[match(OUTPUT$ACTIVITES, output_det$ACTIV_DET)]
OUTPUT$LIV_LOTS_COMPLETS[is.na(OUTPUT$LIV_LOTS_COMPLETS)] <- 0
OUTPUT[15, "LIV_LOTS_COMPLETS"] <- sum(OUTPUT[1:14,"LIV_LOTS_COMPLETS"])

OUTPUT$EMPLOIS <- round(OUTPUT$EMPLOIS, 0)
OUTPUT$LIV_AN <- round(OUTPUT$LIV_AN, 1)
OUTPUT$LIV_MOINS_1KG <- round(OUTPUT$LIV_MOINS_1KG, 1)
OUTPUT$LIV_COLIS_MOINS_3KG <- round(OUTPUT$LIV_COLIS_MOINS_3KG, 1)
OUTPUT$LIV_COLIS_MOINS_30KG <- round(OUTPUT$LIV_COLIS_MOINS_30KG, 1)
OUTPUT$LIV_COLIS_MOINS_100KG <- round(OUTPUT$LIV_COLIS_MOINS_100KG, 1)
OUTPUT$LIV_COLIS_PLUS_100KG <- round(OUTPUT$LIV_COLIS_PLUS_100KG, 1)
OUTPUT$LIV_PALETTES <- round(OUTPUT$LIV_PALETTES, 1)
OUTPUT$LIV_AUTRE <- round(OUTPUT$LIV_AUTRE, 1)
OUTPUT$LIV_LOTS_COMPLETS <- round(OUTPUT$LIV_LOTS_COMPLETS, 1)

write.xlsx(OUTPUT, file = paste0(chemin, "output/", fichier_b2b,"_output.xlsx"), 
           sheetName = "OUTPUT", colClasses = "numeric", append = TRUE)

rm(list = ls(all.names = TRUE))



