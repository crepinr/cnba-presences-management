#!/usr/bin/env bash
set -euo pipefail

CONFIG_FILE="attendance_config.yaml"
ENV_FILE=".env"
UPLOAD="0"
ALLOW_UNMATCHED="0"
OUTPUT_SQL=""
DB_HOST=""
DB_USER=""
DB_NAME=""

usage() {
  cat <<'EOF'
Usage:
  ./run_attendance_pipeline.sh [options]

Options:
  --config <file>       Fichier YAML. Défaut: attendance_config.yaml
  --env-file <file>     Fichier .env. Défaut: .env
  --output <file.sql>   Fichier SQL unique à générer
  --upload              Upload phpMyAdmin après génération SQL
  --no-upload           Ne pas uploader. Défaut
  --allow-unmatched     Autorise les nageurs non reconnus
  --db-host <host>      DB locale pour contrôle métier. Défaut: CNBA_DB_HOST ou localhost
  --db-user <user>      DB locale pour contrôle métier. Défaut: CNBA_DB_USER ou root
  --db-name <db>        DB locale pour contrôle métier. Défaut: CNBA_DB_NAME ou pcxa_cnba
  -h, --help            Aide

Le pipeline traite uniquement les groupes YAML avec active: true.
Les identifiants DB/phpMyAdmin peuvent être définis dans le fichier .env.
EOF
}

load_env_file() {
  local file="$1"

  if [[ ! -f "$file" ]]; then
    if [[ "$file" == ".env" ]]; then
      return 0
    fi
    echo "Fichier .env introuvable: $file" >&2
    exit 1
  fi

  # Charge un .env simple au format CLE=valeur.
  # Les variables déjà exportées peuvent être surchargées par le fichier.
  set -a
  # shellcheck disable=SC1090
  source "$file"
  set +a
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --config)
      CONFIG_FILE="$2"
      shift 2
      ;;
    --env-file)
      ENV_FILE="$2"
      shift 2
      ;;
    --output)
      OUTPUT_SQL="$2"
      shift 2
      ;;
    --upload)
      UPLOAD="1"
      shift
      ;;
    --no-upload)
      UPLOAD="0"
      shift
      ;;
    --allow-unmatched)
      ALLOW_UNMATCHED="1"
      shift
      ;;
    --db-host)
      DB_HOST="$2"
      shift 2
      ;;
    --db-user)
      DB_USER="$2"
      shift 2
      ;;
    --db-name)
      DB_NAME="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Option inconnue: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ ! -f "$CONFIG_FILE" ]]; then
  echo "Fichier YAML introuvable: $CONFIG_FILE" >&2
  exit 1
fi

load_env_file "$ENV_FILE"

DB_HOST="${DB_HOST:-${CNBA_DB_HOST:-localhost}}"
DB_USER="${DB_USER:-${CNBA_DB_USER:-root}}"
DB_NAME="${DB_NAME:-${CNBA_DB_NAME:-pcxa_cnba}}"

export CNBA_CONFIG_YAML="$CONFIG_FILE"

ACTIVE_GROUPS=()
while IFS= read -r GROUP_LABEL; do
  [[ -n "$GROUP_LABEL" ]] && ACTIVE_GROUPS+=("$GROUP_LABEL")
done < <(python3.11 attendance_yaml_config.py active-labels --config "$CONFIG_FILE")

if [[ ${#ACTIVE_GROUPS[@]} -eq 0 ]]; then
  echo "Aucun groupe actif dans le YAML. Mettre active: true sur au moins un groupe." >&2
  exit 1
fi

SEASON_CODE="$(python3.11 - <<'PY'
import os
import attendance_yaml_config
data = attendance_yaml_config.load_yaml(os.environ["CNBA_CONFIG_YAML"])
season = data.get("season", {}) or {}
print(season.get("code") or f"{season.get('start_year')}-{season.get('end_year')}")
PY
)"

TIMESTAMP="$(date '+%Y%m%d_%H%M%S')"

if [[ -z "$OUTPUT_SQL" ]]; then
  OUTPUT_SQL="output/sql/attendance_${SEASON_CODE}_${TIMESTAMP}.sql"
fi

mkdir -p "$(dirname "$OUTPUT_SQL")"

echo "============================================================"
echo "Pipeline Presence Analytics"
echo "YAML      : $CONFIG_FILE"
echo "ENV       : $ENV_FILE"
echo "Saison    : $SEASON_CODE"
echo "Groupes   : ${ACTIVE_GROUPS[*]}"
echo "SQL final : $OUTPUT_SQL"
echo "============================================================"

for GROUP in "${ACTIVE_GROUPS[@]}"; do
  echo
  echo "############################################################"
  echo "# Groupe: $GROUP"
  echo "############################################################"

  python3.11 00_download_data.py --group "$GROUP"
  python3.11 01_import_data.py --group "$GROUP"
  python3.11 02_clean_up.py --group "$GROUP"
done

GROUPS_CSV="$(IFS=, ; echo "${ACTIVE_GROUPS[*]}")"
GROUP_ID_MAP="$(python3.11 attendance_yaml_config.py group-id-map --config "$CONFIG_FILE" --groups "$GROUPS_CSV")"

EXPORT_ARGS=(
  "05_export_attendance_sql.py"
  "--groups" "$GROUPS_CSV"
  "--season-code" "$SEASON_CODE"
  "--output" "$OUTPUT_SQL"
  "--db-host" "$DB_HOST"
  "--db-user" "$DB_USER"
  "--db-name" "$DB_NAME"
)

if [[ -n "$GROUP_ID_MAP" ]]; then
  EXPORT_ARGS+=("--group-id-map" "$GROUP_ID_MAP")
fi

if [[ "$ALLOW_UNMATCHED" == "1" ]]; then
  EXPORT_ARGS+=("--allow-unmatched")
fi

python3.11 "${EXPORT_ARGS[@]}"

if [[ ! -s "$OUTPUT_SQL" ]]; then
  echo "Le fichier SQL attendu n'existe pas ou est vide: $OUTPUT_SQL" >&2
  exit 1
fi

if [[ "$UPLOAD" == "1" ]]; then
  python3.11 06_selenium_upload_sql.py --file "$OUTPUT_SQL" --env-file "$ENV_FILE"
else
  echo
  echo "Upload désactivé. Fichier SQL prêt: $OUTPUT_SQL"
fi
