# Kamerix – Parcours des paramètres (essais manuels du lot 5.6)

Commandes dans la fenêtre Exécution (Ctrl+G), puis Entrée.

## 1. Base neuve et assistant
1. Tests d'abord : `TestSettings` (41/41), `TestDocuments` (49/49).
2. Supprimer `Database\Kamerixdb.accdb` et le dossier `Ressources` s'il existe ; lancer `TestDatabaseCreation`.
3. `StartKamerix` → créer l'administrateur → noter le code de secours.
**Attendu** : « Bienvenue dans Kamerix », seul l'onglet **Entreprise** est visible, champs vides (pied de page :
« Merci de votre confiance. »), boutons **Enregistrer** et **Plus tard**.

## 2. Report de la configuration
1. Cliquer **Plus tard**. **Attendu** : message « Configuration reportée… ».
2. `OpenScreen "DECONNEXION"` puis `StartKamerix` (se reconnecter). **Attendu** : l'assistant se rouvre.

## 3. Contrôles de saisie
Dans l'assistant, cliquer **Enregistrer** avec : nom vide → message « Le nom de l'entreprise est obligatoire. » ;
téléphone `12ab` → « Le téléphone est invalide… » ; e-mail `contact@kamerix` → « L'adresse e-mail est invalide. »

## 4. Configuration valide avec logo
Nom **TS AUTO PRO**, adresse, téléphone, e-mail, NIU, RCCM (en minuscules, pour voir la mise en majuscules), pied de page ;
**Parcourir…** → choisir une image **PNG** → l'aperçu s'affiche avec le nom du fichier → **Enregistrer** → Oui.
**Attendu** : « Entreprise configurée… » ; un fichier `Ressources\logo_….png` à côté du classeur.

## 5. Écran PARAMETRES – `OpenScreen "PARAMETRES"`
**Attendu** : 3 onglets ; Entreprise rempli (NIU et RCCM en majuscules, logo affiché) ; Vente : 19,25 / Espèces, Mobile Money,
Virement bancaire (une ligne chacun) / FCFA / 0 ; Sécurité : 5 / 15.
1. Onglet Vente : TVA `abc` → Enregistrer → message et retour sur l'onglet Vente. Remettre 19,25.
2. Ajouter une ligne `Chèque` aux modes ; Sécurité : 3 et 30 → Enregistrer → « Paramètres enregistrés. »
3. Rouvrir PARAMETRES : les valeurs sont conservées. Remettre 5 et 15 si vous le souhaitez.

## 6. Documents débloqués et verrou de la devise
1. Référentiel, tiers, un achat, une vente, un encaissement (comme PARCOURS_VENTE.md étapes 1 à 5) :
   « Imprimer le reçu ? » est proposé ; le PDF porte **TS AUTO PRO**, le **logo**, NIU et RCCM ; le mode **Chèque** est proposé à l'encaissement.
2. `OpenScreen "PARAMETRES"` → onglet Vente. **Attendu** : devise et décimales **grisées**, message orange
   « Devise et décimales verrouillées : des ventes sont déjà enregistrées. »

## 7. Rôles
1. `OpenScreen "COMPTES"` → créer **vendeur1** (rôle Vendeur) ; `OpenScreen "DECONNEXION"` ; se connecter en vendeur1.
2. `OpenScreen "PARAMETRES"`. **Attendu** : « Seul un administrateur peut modifier les paramètres de Kamerix. »

## Résultat à me renvoyer
« Tout OK », ou le numéro de l'étape en écart avec le message exact (ou une capture).
