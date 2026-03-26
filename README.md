# 📊 Analyse Comptable post-migration

Application Streamlit de comparaison de fichiers comptables (Grand Livre & Balance Auxiliaire Fournisseurs).

## 🚀 Déploiement sur Render

### Étape 1 — Préparer le dépôt GitHub

1. Créez un nouveau dépôt GitHub (public ou privé)
2. Placez ces fichiers à la racine du dépôt :
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - `.gitignore`
3. Poussez vers GitHub :
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/VOTRE_USERNAME/VOTRE_REPO.git
   git push -u origin main
   ```

### Étape 2 — Créer le service sur Render

1. Connectez-vous sur [render.com](https://render.com)
2. Cliquez **"New +"** → **"Web Service"**
3. Connectez votre compte GitHub et sélectionnez le dépôt
4. Render détectera automatiquement le fichier `render.yaml`
5. Vérifiez les paramètres :
   - **Runtime** : Python 3
   - **Build Command** : `pip install -r requirements.txt`
   - **Start Command** : `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`
6. Cliquez **"Create Web Service"**

### Étape 3 — Accéder à l'application

Après le build (2-3 minutes), votre application sera disponible à l'URL :
```
https://analyse-comptable.onrender.com
```
(le nom dépend du nom de service choisi)

---

## 🛠 Lancement local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📦 Dépendances

- `streamlit` — Interface web
- `pandas` — Traitement des données
- `openpyxl` — Export Excel
