# Application Compte Rendu Essais GE

Application web Flask simple pour que vos collègues remplissent les comptes rendus d'essais des groupes électrogènes, avec consultation admin et export Excel.

## Fonctionnalités

- Formulaire public (basé sur votre HTML) pour saisir les rapports
- Stockage en base SQLite
- Support optionnel des photos (2 photos)
- Interface admin avec login
- Liste des rapports + vue détaillée
- Export complet en Excel (.xlsx)

## Installation

```bash
cd ge_report_app
python -m venv venv
source venv/bin/activate   # ou venv\Scripts\activate sur Windows
pip install -r requirements.txt
python app.py
```

Ouvrir http://localhost:5000 pour le formulaire.

Admin : http://localhost:5000/login  
**Identifiants par défaut : admin / admin123**

⚠️ **Changez le mot de passe** dans `app.py` (variable ADMIN_PASSWORD_HASH). Utilisez :
```python
from werkzeug.security import generate_password_hash
print(generate_password_hash('votre_mot_de_passe'))
```

## Notes importantes

1. **Sécurité** : Changez `app.secret_key` et le mot de passe admin avant mise en production. Ajoutez HTTPS.

3. **Hébergement** : Peut tourner sur n'importe quel serveur Python (VPS, Railway, Render, PythonAnywhere, etc.). SQLite suffit pour un usage interne.

4. **Photos** : Stockées dans le dossier `uploads/` (crée automatiquement). Accessibles seulement aux admins connectés.

5. **Export Excel** : Bouton dans l'admin. Tous les champs du JSON sont aplatis en colonnes.

## Améliorations possibles

- Plusieurs comptes admin
- Filtres par date/bâtiment
- Suppression de rapports
- Email de notification
- Déploiement Docker

Pour toute question, adaptez le code selon vos besoins.
