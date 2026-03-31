# Webinar to CRM

Script Python qui lit un fichier Excel d'export **LiveStorm** et met à jour automatiquement les fiches leads correspondantes dans **noCRM.io** en enrichissant leur description avec les informations de participation. Parfait pour les exports natifs de LiveStorm.

---

## Fonctionnement

1. Lecture du fichier `.xlsx` présent dans `webinar_to_crm/data_xlsx/`
2. Suppression des colonnes inutiles (données techniques, UTM, navigateur, etc.)
3. Pour chaque participant, recherche du lead correspondant dans noCRM via son adresse email
4. Mise à jour de la description du lead avec une mention de participation au webinar :
   - ✅ **Présent** → `"{Prénom} {Nom} a participé au Webinar xxx avec un taux de présence de {taux}"`
   - ❌ **Absent** → `"{Prénom} {Nom} s'est inscrit mais n'a pas participé au Webinar xxx"`
   - ❓ **Autre** → `"{Prénom} {Nom} s'est inscrit au Webinar xxx"`
5. Sauvegarde du fichier nettoyé dans `data_cleaned/data_cleaned.xlsx`

La description est mise à jour de façon ciblée : le script localise le **bloc contact** correspondant à l'email dans la description existante et y ajoute ou complète la ligne `Source :`, sans écraser le reste.

---

## Prérequis

- Python 3.8+
- Un compte **noCRM.io** avec accès API

### Dépendances

```bash
pip install openpyxl python-dotenv requests
```

---

## Configuration

Créer un fichier `.env` à la racine du projet :

```env
NOCRM_API_KEY=''
NOCRM_SUBDOMAIN=''
```

| Variable          | Description                                              |
|-------------------|----------------------------------------------------------|
| `NOCRM_API_KEY`   | Clé API noCRM (disponible dans les paramètres du compte) |
| `NOCRM_SUBDOMAIN` | Sous-domaine noCRM (ex: `monentreprise` pour `monentreprise.nocrm.io`) |

---

## Structure du projet

```
.
├── webinar_to_crm/
│   └── data_xlsx/          # Déposer ici le fichier Excel exporté du webinar (un seul fichier)
├── data_cleaned/
│   └── data_cleaned.xlsx   # Fichier Excel nettoyé (généré automatiquement)
├── main.py
├── .env
└── README.md
```

---

## Utilisation

1. Déposer le fichier `.xlsx` d'export webinar dans `webinar_to_crm/data_xlsx/`
2. S'assurer que le fichier `.env` est correctement renseigné
3. Lancer le script :

```bash
python main.py
```

---

## Format attendu du fichier Excel

Le fichier doit contenir les colonnes suivantes (les autres sont supprimées automatiquement) :

| Colonne        | Description                          |
|----------------|--------------------------------------|
| Email          | Adresse email du participant         |
| Prénom         | Prénom                               |
| Nom            | Nom de famille                       |
| Présent        | `True` / `False`                     |
| Taux de présence | Pourcentage de présence au webinar |
| Company        | Entreprise                           |
| Job title      | Intitulé de poste                    |
| Phone number   | Numéro de téléphone                  |

---

## Format attendu des descriptions noCRM

Le script s'appuie sur la structure des descriptions noCRM pour localiser le bon bloc contact. Chaque bloc est séparé par une ligne de tirets (`----------`) et doit contenir une ligne `Email :` pour être identifié.

Exemple de description avec plusieurs contacts :

```
----------
Nom : John Doe
Fonction : Product Manager
Téléphone : +33 x xx xx xx xx
Email : john.doe@xxxx.com
Source : https://www.linkedin.com/in/xxxxxxxxxxxx/
----------
Nom : Jane Doe
Fonction : Owner
Téléphone : +33 x xx xx xx xx
Email : jane.doe@xxxx.com
Source : https://www.linkedin.com/in/xxxxxxxxxxxx/
----------
```

Le script identifie le bloc dont la ligne `Email :` correspond à l'adresse du participant LiveStorm, puis complète la ligne `Source :` :

```
Source : https://www.linkedin.com/in/xxxxxxxxxxxx/, John Doe a participé au Webinar xxx avec un taux de présence de 80%
```

Si la ligne `Source` est absente du bloc, elle sera ajoutée automatiquement.

---

## Notes

- Si aucun lead n'est trouvé pour un email donné, le script l'indique en console et passe au suivant.
- Si aucun bloc contact correspondant à l'email n'est trouvé dans la description, le lead n'est pas modifié.
- Le fichier Excel nettoyé est toujours sauvegardé, même en cas d'erreurs API partielles.
