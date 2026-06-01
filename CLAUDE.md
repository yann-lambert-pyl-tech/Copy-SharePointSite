# Convention projet — Stream Coding

Tout le code de ce projet suit la logique **« stream coding »**.

## Principe

Le code est organisé en **streams** : des sections nommées, clairement délimitées, chacune **autonome**, **commentée en tête** et à **responsabilité unique**. On lit et on maintient le code *stream par stream*.

- **PowerShell** : blocs `#region <Nom du stream>` / `#endregion`.
- **Autres langages** : séparateurs de blocs équivalents (commentaires de section bien visibles).

## Streams standards

Ordonner les streams ainsi (adapter selon le besoin, mais garder l'ordre logique) :

1. **Config / Paramètres** — paramètres d'entrée, validation, constantes
2. **Pré-requis** — vérification des dépendances / modules
3. **Journalisation (Log)** — fonction de log unifiée (fichier + console)
4. **Interface / Suivi (UI)** — barres de progression, rapports d'avancement
5. **Authentification** — connexion, secrets, contexte de sécurité
6. **Logique métier** — un stream par responsabilité fonctionnelle
7. **Vérification & rapport** — contrôles post-exécution, résumé final
8. **Orchestration (Main)** — enchaînement des streams, gestion d'erreur globale

## Standards obligatoires dans chaque script

- **Log** : journal horodaté (fichier) + sortie console lisible
- **Authentification** : explicite, jamais de secret en dur
- **UI de suivi** : l'utilisateur doit voir l'avancement (progress + résumé)
- **Gestion d'erreurs** : try/catch, retry avec back-off pour les appels réseau
- **En-tête** : bloc de documentation / aide en début de fichier

## Référence

Voir `Copy-SharePointSite.ps1` comme implémentation de référence de cette convention.
