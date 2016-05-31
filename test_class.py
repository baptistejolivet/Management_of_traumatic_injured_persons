import xlrd


class Donnees_admin:
    
    def __init__(self,donnee_admin):
        """Donnees administratives du patient"""
        self.ref = donnee_admin[0]
        self.date_saisie = donnee_admin[1]
        self.remplissage = donnee_admin[2]
        self.date_entree = donnee_admin[3]
        self.date_sortie = donnee_admin[4]

    def __str__(self):

        return "Ref :{0}, Date saisie :{1}, Remplissage :{2}, Date entree :{3}, Date sortie :{4}".format(self.ref,self.date_saisie,self.remplissage,self.date_entree,self.date_sortie)


class Donnees_patient:
    
    def __init__(self,donnee_patient):
        """Donnees de base du patient"""
        self.initiales = donnee_patient[0]
        self.date_naissance = donnee_patient[1]
        self.age = donnee_patient[2]
        self.sexe = donnee_patient[3]
        self.poids = donnee_patient[4]
        self.taille = donnee_patient[5]
        self.BMI = donnee_patient[6]
        self.sit_pro = donnee_patient[7]
        self.acti_soc = donnee_patient[8]
        self.detail_act = donnee_patient[9]
        self.vivant = donnee_patient[10]
        self.niveau_et = donnee_patient[11]
        self.antecedents_ASA_PS = donnee_patient[12]
        self.antecedents_grossesse = donnee_patient[13]
        self.trait_anticoag = donnee_patient[14]
        self.trait_antiagr = donnee_patient[15]
        

    def __str__(self):

        return "Initiales: {0}, Date naissance: {1}, Age: {2}, Sexe: {3}, Poids: {4}, Taille: {5}, BMI: {6}, Situation pro: {7}, Activite socio: {8}, Detail activite: {9}, Vivant: {10}, Niveau etudes: {11}, Antecedents ASA_PS: {12}, Antecedents Grossesses: {13}, Traitement anticoag: {14}, Traitement antiagr: {15}".format(self.ref,self.date_saisie,self.remplissage,self.date_entree,self.date_sortie)


class Contexte_accident:
    
    def __init__(self,cont_acc):
        """Donnees administratives du patient"""
        self.Mecanisme = cont_acc[0]
        self.Intentionnalite = cont_acc[1]
        self.Activite_trauma = cont_acc[2]
        self.Detail_activite = cont_acc[3]
        self.Lieu = cont_acc[4]
        self.Type_lieu = cont_acc[5]
        self.Detail_Type_lieu = cont_acc[6]
        self.Ejection_vehicule = cont_acc[7]
        self.Passager_decede_mm_vehicule = cont_acc[8]
        self.AGV = cont_acc[9]
        self.Incarceration = cont_acc[10]
        self.Duree_incarceration = cont_acc[11]
        self.Cin_NSP = cont_acc[12]
        self.Ecrase_projete = cont_acc[13]
        self.Chute = cont_acc[14]
        self.Blast = cont_acc[15]
        self.Isch_du_membre = cont_acc[16]
        self.Amputation = cont_acc[17]
        self.Bassin = cont_acc[18]
        

    def __str__(self):

        return "Mecanisme :{0}, Intentionnalite :{1}, Actitive_trauma :{2}, Detail_activite :{3}, Lieu :{4}, Type_lieu :{5}, Detail_Type_lieu :{6}, Ejecction_vehicule :{7}, Passager decede mm vehic :{8}, AGV :{9}, Incarceration : {10}, Duree incarceration : {11}, Cin_NSP : {12}, Ecrase_projete : {13}, Chute : {14}, Blast : {15}, Isch_du_membre : {16}, Amputation : {17}, Bassin : {18}".format(self.Mecanisme, self.Intentionnalite, self.Activite_trauma, self.Detail_activite, self.Lieu, self.Type_lieu,self.Detail_Type_lieu, self.Ejection_vehicule, self.Passager_decede_mm_vehicule, self.AGV, self.Incarceration, self.Duree_incarceration, self.Cin_NSP, self.Ecrase_projete, self.Chute, self.Blast, self.Isch_du_membre, self.Amputation, self.Bassin)



class Donnees_exam_SAMU:
    
    def __init__(self,donnee_ex_SAMU):
        """Donnees examen a l arrivee du SAMU"""
        self.Glasgow_init = donnee_ex_SAMU[0]
        self.Glasgow_mot_init = donnee_ex_SAMU[1]
        self.PAS_min = donnee_ex_SAMU[2]
        self.PAD_min = donnee_ex_SAMU[3]
        self.FC_max = donnee_ex_SAMU[4]
        self.ACR1 = donnee_ex_SAMU[5]
        self.Lactactes_prehosp = donnee_ex_SAMU[6]
        self.Hemocue_init = donnee_ex_SAMU[7]
        self.SpO2_min = donnee_ex_SAMU[8]
        self.Mydriase = donnee_ex_SAMU[9]
        self.Signe_localisation = donnee_ex_SAMU[10]
        self.AZ = donnee_ex_SAMU[11]
        self.Douleur_palp_rachis = donnee_ex_SAMU[12]
        self.BB = donnee_ex_SAMU[13]

    def __str__(self):

        return "Glasgow_init :{0}, Glasgow_mot_init :{1}, PAS_min:{2}, PAD_min :{3}, FC_max :{4}, ACR1:{5}, Lactates_prehosp :{6}, Hemocue_init :{7}, SpO2_min :{8}, Mydriase :{9}, Signe_localisation : {10}, AZ : {11}, Douleur_palp_rachis : {12}, BB : {13}".format(self.Glasgow_init,self.Glasgow_mot_init,self.PAS_min,self.PAD_min,self.FC_max,self.ACR1,self.Lactactes_prehosp,self.Hemocue_init,self.SpO2_min,self.Mydriase,self.Signe_localisation,self.AZ,self.Douleur_palp_rachis,self.BB)


class Mesure_therap_SAMU:
    
    def __init__(self,mes_th_SAMU):
        """Mesure therapeutiques SAMU"""
        self.Mannitol_SSH = mes_th_SAMU[0]
        self.Reg_myd_osmot = mes_th_SAMU[1]
        self.Rempl_tot_cristalloide = mes_th_SAMU[2]
        self.Rempl_total_colloide = mes_th_SAMU[3]
        self.Catecholamines = mes_th_SAMU[4]
        self.IOT_SMUR = mes_th_SAMU[5]

    def __str__(self):

        return "Mannitol_SSH :{0}, Reg_myd_osmot :{1}, Rempl_tot_cristalloide:{2}, Rempl_total_colloide :{3}, Catecholamines :{4}, IOT_SMUR :{5}".format(self.Mannitol_SSH, self.Reg_myd_osmot, self.Rempl_tot_cristalloide, self.Rempl_total_colloide, self.Catecholamines, self.IOT_SMUR)


class Donnees_transfert:
    
    def __init__(self,donnee_transf):
        """Donnees de transfert"""
        self.Mode = donnee_transf[0]
        self.Centre = donnee_transf[1]
        self.Auteur = donnee_transf[2]
        self.Discordance = donnee_transf[3]
        self.Annonce_instable = donnee_transf[4]
        self.Origine_patient = donnee_transf[5]
        self.depart_SAMU = donnee_transf[6]
        self.delai_alarme_sur_lieux_SMUR = donnee_transf[7]
        self.delai_sur_lieux_hopital = donnee_transf[8]
        self.activation_procedure_choc = donnee_transf[9]
        self.periode_travail = donnee_transf[10]

    def __str__(self):

        return "Mode :{0}, Centre :{1}, Auteur :{2}, Discordance :{3}, Annonce instable :{4}, Origine patient :{5}, Depart SAMU  :{6}, Delai alarme  sur les lieux SMUR :{7}, Delai sur lieux Hopital :{8}, Activation prodecure choc :{9}, Periode travail :{10}".format(self.Mode,self.Centre,self.Auteur,self.Discordance,self.Annonce_instable,self.Origine_patient,self.depart_SAMU,self.delai_alarme_sur_lieux_SMUR,self.delai_sur_lieux_hopital,self.activation_procedure_choc,self.periode_travail)


class Examen_clinique:

    def __init__(self,exam_clin):
        """Donnees de transfert"""
        self.Hemocue_arr_hop = exam_clin[0]
        self.Glasgow = exam_clin[1]
        self.Glasgow_mot = exam_clin[2]
        self.Anomalie_pupille = exam_clin[3]
        self.FC = exam_clin[4]
        self.PAS = exam_clin[5]
        self.PAD = exam_clin[6]
        self.SpO2 = exam_clin[7]
        self.Temp = exam_clin[8]
        self.Radio_poumon = exam_clin[9]
        self.Radio_bassin = exam_clin[10]
        self.Fract_ouv = exam_clin[11]
        self.Temp_min = exam_clin[12]
        self.FAST = exam_clin[13]
        self.DTC_IP_max1 = exam_clin[14]
        self.pH = exam_clin[15]
        self.Exces_de_base = exam_clin[16]
        self.Lactates = exam_clin[17]
        self.Lactates_H2 = exam_clin[18]
        self.Lactates_sup_H2 = exam_clin[19]
        self.pCO2 = exam_clin[20]
        self.PaO2 = exam_clin[21]
        self.Ventilation_FiO2 = exam_clin[22]
        self.FiO2 =exam_clin[23]
        self.Hb = exam_clin[24]
        self.Plaquettes = exam_clin[25]
        self.TP_pourc = exam_clin[26]
        self.TP_INR = exam_clin[27]
        self.Fibriogene = exam_clin[28]
        self.Creatinine = exam_clin[29]
        self.Bicar = exam_clin[30]
        self.Tropo = exam_clin[31]
        self.Alcool = exam_clin[32]
        
        


class Decision_med_therap:

    def __init__(self,dec_med_th):
        """Donnees de transfert"""
        self.Drainage_thor = dec_med_th[0]
        self.Bloc_direct = dec_med_th[1]
        self.Delai_dep_scan_ou_bloc = dec_med_th[2]
        self.KTV_pose_avant_TDM = dec_med_th[3]
        self.Derniere_PAS_avant_scan_ou_bloc = dec_med_th[4]
        self.Derniere_PAD_avant_scan_bloc = dec_med_th[5]
        self.Dose_NAD_dep_scan = dec_med_th[6]
        self.Nb_CGR_dechoc = dec_med_th[7]
        self.Nb_PCF_dechoc = dec_med_th[8]
        self.Nb_CP_dechoc = dec_med_th[9]
        self.Acide_tranexamique = dec_med_th[10]
        self.Fibriogene = dec_med_th[11]
        self.Intervention = dec_med_th[12]
        
class Blanche:

    def __init__(self,blanche):
        """Donnees de transfert"""
        self.IGS_II = blanche[0]
        self.SOFA_24h = blanche[1]
        self.ATB_J1 = blanche[2]
        self.Choc_hemo = blanche[3]
        self.Or_saignement = blanche[4]
        self.Prem_type_act_th_essentielle = blanche[5]
        self.Delai_adm_geste = blanche[6]
        self.Tech_dam_contr_utilisee = blanche[7]
        self.ACR2 = blanche[8]
        self.Delai_intro_catechol = blanche[9]
        self.Catechol_max = blanche[10]
        self.CGR_6h = blanche[11]
        self.PCF_6h = blanche[12]
        self.Plaquettes_6h = blanche[13]
        self.Fibrinogene_6h = blanche[14]
        self.Cristalloides_24h = blanche[15]
        self.Colloides_24h = blanche[16]
        self.CGR_24h = blanche[17]
        self.PCF_24h = blanche[18]
        self.Plaquettes_24h = blanche[19]
        self.Fibrinogene_24h = blanche[20]
        self.Novoseven = blanche[21]
        self.Ac_tranexamique = blanche[22]

    

class Prise_en_charge_trauma_cranien:
    """Prise en charge Trauma cranien"""
    def __init__(self,t_cranien):
        """Donnees de transfert"""
        self.Trauma_cranien = t_cranien[0]
        self.Osmotherapie = t_cranien[1]
        self.PIC = t_cranien[2]
        self.Delai_PIC = t_cranien[3]
        self.DVE = t_cranien[4]
        self.Delai_DVE = t_cranien[5]
        self.DTC_IP_max2 = t_cranien[6]
        self.Cranectiomie = t_cranien[7]

    def __str__(self):

        return "Trauma cranien :{0}, Osmotherapie :{1}, PIC :{2}, Delai PIC:{3}, DVE :{4}, Delai DVE :{5}, DTC IP max :{6}, Cranectiomie :{7}".format(self.Trauma_cranien,self.Osmotherapie,self.PIC,self.Delai_PIC,self.DVE,self.Delai_DVE,self.DTC_IP_max2,self.Cranectiomie)



class Prise_en_charge_trauma_rachis:
    """Prise en charge Trauma rachis"""
    def __init__(self,t_rachis):
        """Prise en charge Trauma rachis"""
        self.Trauma_rachis = t_rachis[0]
        self.rachis_neuro = t_rachis[1]
        self.lesion_compl = t_rachis[2]
        self.niveau_medullaire = t_rachis[3]
        self.ASIA_mot = t_rachis[4]
        self.pH_min = t_rachis[5]
        self.Lact_max = t_rachis[6]
        self.creatinine_max = t_rachis[7]
        self.Ca_ionise_min = t_rachis[8]
        self.PF_min = t_rachis[9]
        self.BE_max = t_rachis[10]
        self.Bicar_min = t_rachis[11]
        self.TP_min_pourc = t_rachis[12]
        self.TP_min_INR = t_rachis[13]
        self.Fibrinogene_min = t_rachis[14]
        self.HB_min = t_rachis[15]
        self.Plaquettes_min = t_rachis[16]

    def __str__(self):

        return "Trauma rachis :{0}, Rachis neuro :{1}, Lesion complete :{2}, Niveau medullaire:{3}, Score ASIA moteur:{4}, pH min:{5}, Lactates max :{6}, Creatinine max :{7}, Ca ionise min:{8}, PF min: {9}, BE max:{10}, Bicar min: {11}, TP min pourc:{12}, TP min INR: {13}, Fibrinogene min : {14}, HB min:{15}, Plaquettes min : {16}".format(self.Trauma_rachis,self.rachis_neuro,self.lesion_compl,self.niveau_medullaire,self.ASIA_mot,self.pH_min,self.Lact_max,self.creatinine_max,self.Ca_ionise_min,self.PF_min,self.BE_max,self.Bicar_min,self.TP_min_pourc,self.TP_min_INR,self.Fibrinogene_min,self.HB_min,self.Plaquettes_min)


class Outcomes_et_complications:
    
    def __init__(self,outcomes):
        """Outcomes et complications"""
        self.Duree_rea = outcomes[0]
        self.Deces = outcomes[1]
        self.Delai_DC = outcomes[2]
        self.Cause_DC = outcomes[3]
        self.Detail_cause_DC = outcomes[4]
        self.Transfert_sec_why = outcomes[5]
        self.Sortie = outcomes[6]
        self.Nb_j_sortant_att_place = outcomes[7]
        self.LATA = outcomes[8]
        self.Nb_intubations = outcomes[9]
        self.HTIC = outcomes[10]
        self.Glasgow_de_sortie = outcomes[11]
        self.Rachis_recup_ASIA = outcomes[12]
        self.Tracheo = outcomes[13]
        self.Nb_j_VM = outcomes[14]
        self.ALI_ARDS = outcomes[15]
        self.RIFLE_max = outcomes[16]
        self.EER = outcomes[17]
        self.Nb_j_EER = outcomes[18]
        self.Nb_j_tot_CGR = outcomes[19]
        self.Nb_tot_PCF = outcomes[20]
        self.Nb_tot_CP = outcomes[21]
        self.SOFA_J2 = outcomes[22]
        self.SOFA_J5 = outcomes[23]
        self.SOFA_J7 = outcomes[24]
        self.Delai_norm_BE = outcomes[25]
        self.Delai_norm_lactates = outcomes[26]
        self.CPK_max = outcomes[27]
        self.Sepsis_inf_cause = outcomes[28]
        self.Sepsis_inf_nb_pneumopath = outcomes[29]
        self.Sepsis_inf_jour1_pneumopath = outcomes[30]
        self.Sepsis_inf_choc_scept = outcomes[31]
        self.Sepsis_inf_acq_BMR = outcomes[32]
        self.NB_j_hop = outcomes[33]
        self.AIS_tete_cou = outcomes[34]
        self.AIS_visage = outcomes[35]
        self.AIS_thorax = outcomes[36]
        self.AIS_organes_abdopelv = outcomes[37]
        self.AIS_mb_ceintpelv = outcomes[38]
        self.AIS_ext = outcomes[39]
        self.ISS = outcomes[40]
        self.Blian_les_AIS = outcomes[41]
        self.Bilan_les_CIM10 = outcomes[42]





class Patient(Donnees_admin,Donnees_patient,Contexte_accident,Donnees_exam_SAMU,Mesure_therap_SAMU,Donnees_transfert,Examen_clinique,Decision_med_therap,Blanche,Prise_en_charge_trauma_cranien,Prise_en_charge_trauma_rachis,Outcomes_et_complications):
    def __init__(self,donnee_admin,donnee_patient,cont_acc,donnee_ex_SAMU,mes_th_SAMU,donnee_transf,exam_clin,dec_med_th,blanche,t_cranien,t_rachis,outcomes):
        Donnees_admin.__init__(self,donnee_admin)
        Donnees_patient.__init__(self,donnee_patient)
        Contexte_accident.__init__(self,cont_acc)
        Donnees_exam_SAMU.__init__(self,donnee_ex_SAMU)
        Mesure_therap_SAMU.__init__(self,mes_th_SAMU)
        Donnees_transfert.__init__(self,donnee_transf)
        Examen_clinique.__init__(self,exam_clin)
        Decision_med_therap.__init__(self,dec_med_th)
        Blanche.__init__(self,blanche)
        Prise_en_charge_trauma_cranien.__init__(self,t_cranien)
        Prise_en_charge_trauma_rachis.__init__(self,t_rachis)
        Outcomes_et_complications.__init__(self,outcomes)


# ouverture du fichier Excel
wb = xlrd.open_workbook('Base_stratifiee.xlsx')
 
# feuilles dans le classeur
##print wb.sheet_names()

# lecture des données dans la première feuille
sh = wb.sheet_by_name(u'Base')
##for rownum in range(sh.nrows):
##    print sh.row_values(rownum)

# lecture par colonne

RefTB = sh.col_values(0)
RefTB = RefTB[1:]

#print colonne1

colonne2=sh.col_values(1)
#print colonne2
 
# lecture des catégories : 

##line1 = sh.row_values(0)

# choix d un patient en fct de sa ref
##print "Reference du patient"
##ref = input()
##line2 = sh.row_values(RefTB.index(ref)+1)
##
##
##donnee_admin = line2[0:5]
##donnee_patient = line2[5:21]
##cont_acc = line2[21:40]
##donnee_ex_SAMU = line2[40:54]
##mes_th_SAMU = line2[54:60]
##donnee_transf = line2[60:71]
##exam_clin = line2[71:104]
##dec_med_th = line2[104:117]
##blanche = line2[117:140]
##t_cranien = line2[140:148]
##t_rachis = line2[148:165]
##outcomes = line2[165:208]
##
##
##patient = Patient(donnee_admin,donnee_patient,cont_acc,donnee_ex_SAMU,mes_th_SAMU,donnee_transf,exam_clin,dec_med_th,blanche,t_cranien,t_rachis,outcomes)
##print "Centre:",patient.Centre
##print "Mode de transfert:",patient.Mode
##print "Origine:", patient.Origine_patient
##print "Mecanisme:", patient.Mecanisme
##print "Deces:", patient.Deces
##print "age", patient.age
##print "Hb", patient.Hb


## TRI DES PATIENTS EN FCT DE LEUR ETAT (MORT,VIVANT,NSP) ou DISCORDANCE:

Morts = []
Vivants = []
NSP = []

Discordance = []
Origine = []

for i in range(len(RefTB)):

    ref = RefTB[i]
    line = sh.row_values(RefTB.index(ref)+1)

    donnee_admin = line[0:5]
    donnee_patient = line[5:21]
    cont_acc = line[21:40]
    donnee_ex_SAMU = line[40:54]
    mes_th_SAMU = line[54:60]
    donnee_transf = line[60:71]
    exam_clin = line[71:104]
    dec_med_th = line[104:117]
    blanche = line[117:140]
    t_cranien = line[140:148]
    t_rachis = line[148:165]
    outcomes = line[165:208]

    patient = Patient(donnee_admin,donnee_patient,cont_acc,donnee_ex_SAMU,mes_th_SAMU,donnee_transf,exam_clin,dec_med_th,blanche,t_cranien,t_rachis,outcomes)
    
    if patient.Deces == 1.0 :
        Morts = Morts + [ref]

    elif patient.Deces == 0.0 :
        Vivants = Vivants + [ref]

    else :
        NSP = NSP + [ref]

    if patient.Discordance == "Pas de discordance" :
        pass

    elif patient.Discordance == "":
        pass

    elif patient.Discordance == "NA" :
        pass
      
    else :
        Discordance = Discordance + [ref]

    if patient.Origine_patient == "Secondaire" :
        Origine = Origine + [ref]

print "Morts", Morts
print "Vivants", Vivants
print "NSP", NSP
print "Discordance", Discordance
print "Origine", Origine
        
print "Morts", len(Morts)
print "Vivants", len(Vivants)
print "NSP", len(NSP)
print "Discordance", len(Discordance)
print "Origine", len(Origine)

print "Pourcentage Discordance", float(len(Discordance)) / float(len(RefTB))



print ("Intersection entre deux categories")
SET1 = set(input())
SET2 = set(input())
print "Intersection :", SET1 & SET2
print "Taille :", len(SET1 & SET2)
print "Pourcentage :", float(len(SET1 & SET2))/float(len(RefTB))

# Infos complementaires sur les patients concernes :

for i in range(len(SET1 & SET2)):
    line = sh.row_values(RefTB.index(int(list(SET1 & SET2)[i]))+1)
    donnee_admin = line[0:5]
    donnee_patient = line[5:21]
    cont_acc = line[21:40]
    donnee_ex_SAMU = line[40:54]
    mes_th_SAMU = line[54:60]
    donnee_transf = line[60:71]
    exam_clin = line[71:104]
    dec_med_th = line[104:117]
    blanche = line[117:140]
    t_cranien = line[140:148]
    t_rachis = line[148:165]
    outcomes = line[165:208]

    patient = Patient(donnee_admin,donnee_patient,cont_acc,donnee_ex_SAMU,mes_th_SAMU,donnee_transf,exam_clin,dec_med_th,blanche,t_cranien,t_rachis,outcomes)

    print "Mecanisme :",patient.Mecanisme, ", Deces :", patient.Deces, ", Discordance :", patient.Discordance, "Delai :", patient.Delai_adm_geste
    

# Idee generale : arriver à etablir un test d hypothese sur les patients
# Arriver a identifier les patients avec complications et le normaux 


