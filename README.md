# 🎓 Formulaire d'enregistrement des étudiants – VBA Excel

Ce projet vise à développer une **application de gestion des inscriptions étudiantes** à l’aide de **VBA sous Excel**. Il fournit une interface conviviale permettant d’ajouter, consulter et gérer les informations des étudiants à travers un formulaire réparti sur trois feuilles.

---

## 📌 Fonctionnalités principales

- **Interface d’accueil**
  - Message de bienvenue
  - Bouton **"Ajouter un nouvel étudiant"**
  - Bouton **"Voir la liste des étudiants"**

- **Formulaire d’enregistrement**
  - Saisie des informations personnelles :
    - Nom, Prénoms
    - Matricule
    - Date de naissance
    - Âge (calculé automatiquement ou saisi)
    - Option (Économie / Mathématiques)
    - Loisirs (zone de texte)
    - Situation matrimoniale (Célibataire, Marié(e), etc.)
  - Bouton **"Valider l’enregistrement"** ou **"Retour"**

- **Liste des étudiants inscrits**
  - Affichage tabulaire de tous les étudiants enregistrés
  - Bouton **"Quitter"** pour revenir à la page d’accueil

---

## 🛠️ Technologies utilisées

- **Microsoft Excel** (fichier `.xlsm`)
- **VBA (Visual Basic for Applications)** pour la logique et les contrôles
- **Formulaires UserForm** et macros assignées aux boutons

---

## 📁 Structure du fichier Excel

| Feuille             | Rôle                                     |
|---------------------|------------------------------------------|
| `Accueil`           | Interface de navigation principale       |
| `Enregistrement`    | Formulaire de saisie des informations    |
| `Liste_Étudiants`   | Tableau dynamique des étudiants inscrits |

---

## 💡 Suggestions d’amélioration

- ✅ **Validation automatique** des champs obligatoires avant l’enregistrement
- 🔁 **Calcul automatique de l’âge** à partir de la date de naissance
- 🧹 Bouton **"Réinitialiser le formulaire"** pour effacer les champs
- 🔍 Ajout d’une **fonction de recherche** par nom, matricule, ou option
- 📤 Export possible vers un fichier `.csv` ou `.pdf`
- 🔒 **Protection des feuilles** pour éviter les modifications involontaires
- 💾 Enregistrement automatique dans une **base Access** ou fichier externe pour plus de robustesse (optionnel)

---

## 📸 Aperçu de l'application *(ajoutez des captures d'écran si possible)*

> ![Aperçu Accueil](images/accueil.png)  
> ![Aperçu Enregistrement](images/enregistrement.png)  
> ![Aperçu Liste](images/liste_etudiants.png)

---

## ✍️ Auteur

- **Nom** : [Ton Nom ici]
- **Statut** : Élève Ingénieur / Étudiant en Informatique / Data Analyst Junior
- **Langage** : VBA sous Excel
- 📧 Contact : [ton.email@exemple.com]

---

## 📄 Licence

Ce projet est librement utilisable à des fins pédagogiques ou personnelles. Toute amélioration est la bienvenue via fork ou pull request !
