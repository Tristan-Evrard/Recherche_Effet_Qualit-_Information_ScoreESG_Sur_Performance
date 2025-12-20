import pandas as pd 
import statsmodels.api as sm 
import numpy as np 
import os

#Importation des données 
data = pd.read_csv(r"C:\Users\trist\Desktop\Master MBFA GRFV\Méthodologie pour la recherche\Modele_econometrique_CAC40\Data_modeleconom.csv",
                  encoding="latin-1",
                  index_col=0)

# Définir le dossier de destination
dossier_export = r"C:\Users\trist\Desktop\Master MBFA GRFV\Méthodologie pour la recherche\Data"
os.makedirs(dossier_export, exist_ok=True)

# Fonction pour extraire les résultats d'une régression
def extraire_resultats(results, nom_modele):
    """Extrait les statistiques principales d'une régression OLS"""
    
    # Tableau des coefficients
    coef_table = pd.DataFrame({
        'Variable': results.params.index,
        'Coefficient': results.params.values,
        'Std Error': results.bse.values,
        'T-stat': results.tvalues.values,
        'P-value': results.pvalues.values,
        'IC 95% Min': results.conf_int()[0].values,
        'IC 95% Max': results.conf_int()[1].values
    })
    
    # Statistiques du modèle
    stats_modele = pd.DataFrame({
        'Statistique': ['R²', 'R² ajusté', 'F-statistic', 'Prob (F-statistic)', 
                       'Log-Likelihood', 'AIC', 'BIC', 'Nombre observations'],
        'Valeur': [results.rsquared, results.rsquared_adj, results.fvalue, 
                  results.f_pvalue, results.llf, results.aic, results.bic, 
                  results.nobs]
    })
    
    return coef_table, stats_modele

# Dictionnaire pour stocker tous les résultats
tous_resultats = {}

#################################################### MODELE H1 #################################################### 
print('#################################################### MODELE H1 #################################################### ')

X_H1 = data[['Score ESG - %','ESG x Qualité Information','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
Y_H1 = data['ROA - %']
X_H1 = sm.add_constant(X_H1)

model_H1 = sm.OLS(Y_H1, X_H1)
results_H1 = model_H1.fit()
print(results_H1.summary())

coef_H1, stats_H1 = extraire_resultats(results_H1, 'H1')
tous_resultats['H1'] = {'coefficients': coef_H1, 'statistiques': stats_H1}

#################################################### MODELE H1.a ####################################################
print("""
      #################################################### MODELE H1.a #################################################### """)

X_H1_a = data[['Envrionnemental - % ','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
Y_H1_a = data['Disclosure Score - Qualité Information - %']
X_H1_a = sm.add_constant(X_H1_a)

model_H1_a = sm.OLS(Y_H1_a, X_H1_a)
results_H1_a = model_H1_a.fit()
print(results_H1_a.summary())

coef_H1_a, stats_H1_a = extraire_resultats(results_H1_a, 'H1.a')
tous_resultats['H1.a'] = {'coefficients': coef_H1_a, 'statistiques': stats_H1_a}

#################################################### MODELE H1.b ####################################################
print("""
      #################################################### MODELE H1.b #################################################### """)

X_H1_b = data[['Social - %','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
Y_H1_b = data['Disclosure Score - Qualité Information - %']
X_H1_b = sm.add_constant(X_H1_b)

model_H1_b = sm.OLS(Y_H1_b, X_H1_b)
results_H1_b = model_H1_b.fit()
print(results_H1_b.summary())

coef_H1_b, stats_H1_b = extraire_resultats(results_H1_b, 'H1.b')
tous_resultats['H1.b'] = {'coefficients': coef_H1_b, 'statistiques': stats_H1_b}

#################################################### MODELE H1.c ####################################################
print("""
      #################################################### MODELE H1.c #################################################### """)

X_H1_c = data[['Gouvernance - %','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
Y_H1_c = data['Disclosure Score - Qualité Information - %']
X_H1_c = sm.add_constant(X_H1_c)

model_H1_c = sm.OLS(Y_H1_c, X_H1_c)
results_H1_c = model_H1_c.fit()
print(results_H1_c.summary())

coef_H1_c, stats_H1_c = extraire_resultats(results_H1_c, 'H1.c')
tous_resultats['H1.c'] = {'coefficients': coef_H1_c, 'statistiques': stats_H1_c}

#################################################### MODELE H1.d ####################################################
print("""
      #################################################### MODELE H1.d #################################################### """)

X_H1_d = data[['Envrionnemental - % ','Social - %','Gouvernance - %','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
Y_H1_d = data['Disclosure Score - Qualité Information - %']
X_H1_d = sm.add_constant(X_H1_d)

model_H1_d = sm.OLS(Y_H1_d, X_H1_d)
results_H1_d = model_H1_d.fit()
print(results_H1_d.summary())

coef_H1_d, stats_H1_d = extraire_resultats(results_H1_d, 'H1.d')
tous_resultats['H1.d'] = {'coefficients': coef_H1_d, 'statistiques': stats_H1_d}

###############################################################################
# Export en Excel
###############################################################################

# OPTION 1 : Un fichier par modèle
for nom_modele, resultats in tous_resultats.items():
    nom_fichier = f'regression_{nom_modele.replace(".", "_")}.xlsx'
    chemin_fichier = os.path.join(dossier_export, nom_fichier)
    
    with pd.ExcelWriter(chemin_fichier, engine='openpyxl') as writer:
        resultats['coefficients'].to_excel(writer, sheet_name='Coefficients', index=False)
        resultats['statistiques'].to_excel(writer, sheet_name='Statistiques', index=False)
    
    print(f"✓ {nom_fichier} créé")

# OPTION 2 : Un seul fichier avec tous les modèles
chemin_complet = os.path.join(dossier_export, 'regressions_completes.xlsx')
with pd.ExcelWriter(chemin_complet, engine='openpyxl') as writer:
    for nom_modele, resultats in tous_resultats.items():
        # Limiter la longueur du nom de l'onglet (max 31 caractères Excel)
        nom_onglet_coef = f'{nom_modele}_Coef'[:31]
        nom_onglet_stats = f'{nom_modele}_Stats'[:31]
        
        resultats['coefficients'].to_excel(writer, sheet_name=nom_onglet_coef, index=False)
        resultats['statistiques'].to_excel(writer, sheet_name=nom_onglet_stats, index=False)

print(f"\n✓ Fichier complet créé : regressions_completes.xlsx")
print(f"\n✓ Tous les fichiers ont été exportés dans :\n  {dossier_export}")
import pandas as pd
import numpy as np

###############################################################################
# Préparation des données
###############################################################################
data_H1 = data[['ROA - %','Score ESG - %','Disclosure Score - Qualité Information - %','Taille', 'Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]
data_H1_abcd = data[['Disclosure Score - Qualité Information - %', 'Envrionnemental - % ','Social - %','Gouvernance - %','Taille','Dette nette en Mds euros', 'Secteur Secondaire', 'Secteur Tertiaire']]

###############################################################################
# Fonction de calcul de la matrice de corrélation
###############################################################################
def matrice_correlation(
    df: pd.DataFrame,
    methode: str = "pearson",
    drop_na: bool = True
) -> pd.DataFrame:
    """Calcule la matrice de corrélation d'un DataFrame pandas."""
    
    if not isinstance(df, pd.DataFrame):
        raise TypeError("L'entrée doit être un DataFrame pandas")
    
    if methode not in ["pearson", "spearman", "kendall"]:
        raise ValueError("Méthode invalide : pearson, spearman ou kendall")
    
    data_copy = df.copy()
    data_copy = data_copy.select_dtypes(include="number")
    
    if drop_na:
        data_copy = data_copy.dropna()
    
    if data_copy.shape[1] < 2:
        raise ValueError("Le DataFrame doit contenir au moins deux colonnes numériques")
    
    return data_copy.corr(method=methode)

###############################################################################
# Calcul des statistiques complètes
###############################################################################

# TABLEAU RÉCAPITULATIF COMPLET (Moyenne, Écart-type, Min, Max, Taille)
stats_H1 = pd.DataFrame({
    'Variable': data_H1.columns,
    'Moyenne': [np.mean(data_H1[col]) for col in data_H1.columns],
    'Ecart-type': [np.std(data_H1[col]) for col in data_H1.columns],
    'Min': [np.min(data_H1[col]) for col in data_H1.columns],
    'Max': [np.max(data_H1[col]) for col in data_H1.columns],
    'Taille': [data_H1[col].count() for col in data_H1.columns]
})

stats_H1_abcd = pd.DataFrame({
    'Variable': data_H1_abcd.columns,
    'Moyenne': [np.mean(data_H1_abcd[col]) for col in data_H1_abcd.columns],
    'Ecart-type': [np.std(data_H1_abcd[col]) for col in data_H1_abcd.columns],
    'Min': [np.min(data_H1_abcd[col]) for col in data_H1_abcd.columns],
    'Max': [np.max(data_H1_abcd[col]) for col in data_H1_abcd.columns],
    'Taille': [data_H1_abcd[col].count() for col in data_H1_abcd.columns]
})

# Tableaux séparés pour moyennes et écarts-types (si besoin)
moyennes_H1 = stats_H1[['Variable', 'Moyenne']].copy()
moyennes_H1['Dataset'] = 'H1'

moyennes_H1_abcd = stats_H1_abcd[['Variable', 'Moyenne']].copy()
moyennes_H1_abcd['Dataset'] = 'H1_abcd'

moyennes_totales = pd.concat([moyennes_H1, moyennes_H1_abcd], ignore_index=True)

ecarts_H1 = stats_H1[['Variable', 'Ecart-type']].copy()
ecarts_H1['Dataset'] = 'H1'

ecarts_H1_abcd = stats_H1_abcd[['Variable', 'Ecart-type']].copy()
ecarts_H1_abcd['Dataset'] = 'H1_abcd'

ecarts_totaux = pd.concat([ecarts_H1, ecarts_H1_abcd], ignore_index=True)

# MATRICES DE CORRÉLATION
corr_H1 = matrice_correlation(data_H1, methode="pearson")
corr_H1_abcd = matrice_correlation(data_H1_abcd, methode="pearson")

###############################################################################
# Export en Excel
###############################################################################

# Définir le chemin du dossier de destination
import os
dossier_export = r"C:\Users\trist\Desktop\Master MBFA GRFV\Méthodologie pour la recherche\Data"

# Créer le dossier s'il n'existe pas
os.makedirs(dossier_export, exist_ok=True)

# OPTION 1 : Fichiers Excel séparés
moyennes_totales.to_excel(os.path.join(dossier_export, 'moyennes.xlsx'), index=False, engine='openpyxl')
ecarts_totaux.to_excel(os.path.join(dossier_export, 'ecarts_types.xlsx'), index=False, engine='openpyxl')
stats_H1.to_excel(os.path.join(dossier_export, 'statistiques_H1.xlsx'), index=False, engine='openpyxl')
stats_H1_abcd.to_excel(os.path.join(dossier_export, 'statistiques_H1_abcd.xlsx'), index=False, engine='openpyxl')
corr_H1.to_excel(os.path.join(dossier_export, 'correlation_H1.xlsx'), engine='openpyxl')
corr_H1_abcd.to_excel(os.path.join(dossier_export, 'correlation_H1_abcd.xlsx'), engine='openpyxl')

print(f"✓ Fichiers Excel créés avec succès dans :\n  {dossier_export}\n")
print("Fichiers créés :")
print("  - moyennes.xlsx")
print("  - ecarts_types.xlsx")
print("  - statistiques_H1.xlsx (avec Min, Max, Taille)")
print("  - statistiques_H1_abcd.xlsx (avec Min, Max, Taille)")
print("  - correlation_H1.xlsx")
print("  - correlation_H1_abcd.xlsx")

# OPTION 2 : Un seul fichier Excel avec plusieurs onglets
chemin_fichier_complet = os.path.join(dossier_export, 'statistiques_completes.xlsx')
with pd.ExcelWriter(chemin_fichier_complet, engine='openpyxl') as writer:
    moyennes_totales.to_excel(writer, sheet_name='Moyennes', index=False)
    ecarts_totaux.to_excel(writer, sheet_name='Ecarts-types', index=False)
    stats_H1.to_excel(writer, sheet_name='Stats H1', index=False)
    stats_H1_abcd.to_excel(writer, sheet_name='Stats H1_abcd', index=False)
    corr_H1.to_excel(writer, sheet_name='Corrélation H1')
    corr_H1_abcd.to_excel(writer, sheet_name='Corrélation H1_abcd')

print(f"\n✓ Fichier Excel unique créé : statistiques_completes.xlsx")

###############################################################################
# Affichage console (optionnel)
###############################################################################
print("\n" + "="*80)
print("STATISTIQUES DESCRIPTIVES COMPLÈTES")
print("="*80)
print("\n### DATA H1 ###")
print(stats_H1.to_string(index=False))
print("\n### DATA H1_abcd ###")
print(stats_H1_abcd.to_string(index=False))

print("\n" + "="*80)
print("MATRICES DE CORRÉLATION")
print("="*80)
print("\n### DATA H1 ###")
print(corr_H1)
print("\n### DATA H1_abcd ###")
print(corr_H1_abcd)