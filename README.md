# Exemple de connexion SQL depuis Excel avec VBA

Ce dépôt propose un exemple simple de connexion à une base de données SQL depuis Excel à l’aide de VBA, via **ADODB**.  
L’objectif est de montrer comment :

- ouvrir une connexion SQL,
- exécuter une requête `SELECT`,
- récupérer les résultats,
- et les écrire dans une feuille Excel.

---

## ⚠️ Sécurité avant tout

Le fichier fourni est un **exemple pédagogique**.  
Ne mettez **jamais** de mots de passe ou informations sensibles dans un dépôt public.

Dans ce projet, la chaîne de connexion contient des valeurs factices.  
Vous devez :

- soit utiliser un **DSN sécurisé**,
- soit paramétrer les identifiants en dehors du code (fichier de config, variables d’environnement, etc.).

---

## 📂 Contenu

Le classeur `sql-connection-example.xlsm` contient :

- un module `modSqlConnexion` :
  - `GetSqlConnection` : procédure qui ouvre une connexion ADODB
- un module `modImportSql` :
  - `ImporterDonneesDepuisSQL` : exécute une requête SQL et importe les données dans une feuille `SQL_Data`
  - `TesterConnexionSQL` : test simple de connexion

---

## 🧩 Exemple de scénario

- Connexion à une base SQL Server (adaptable à d’autres SGBD)
- Exécution d’une requête du type :

```sql
SELECT TOP 100 *
FROM MaTable
ORDER BY DateCreation DESC;
