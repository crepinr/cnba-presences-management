# Pipeline Presence Analytics

## Configuration

1. Copier l'exemple d'environnement :

```bash
cp .env.example .env
```

2. Compléter `.env` avec les accès DB locale et phpMyAdmin.

3. Compléter `attendance_config.yaml` :
   - `active: true` pour les groupes à traiter ;
   - `active: false` pour les groupes à ignorer ;
   - `google_sheet.key` ou `google_sheet.url` pour chaque groupe.

## Exécution sans upload

```bash
./run_attendance_pipeline.sh --config attendance_config.yaml --env-file .env --no-upload
```

## Exécution avec upload phpMyAdmin

```bash
./run_attendance_pipeline.sh --config attendance_config.yaml --env-file .env --upload
```

## Sortie SQL unique

Même si plusieurs groupes sont actifs dans le YAML, le pipeline génère un seul fichier SQL final.
