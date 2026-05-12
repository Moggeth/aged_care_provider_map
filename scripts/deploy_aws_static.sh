#!/usr/bin/env bash
set -euo pipefail

BUCKET="${BUCKET:-andromeda-aged-care-provider-map-prod}"
DISTRIBUTION_ID="${DISTRIBUTION_ID:-E2Y6OESD3IEA02}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

aws s3 sync "$ROOT_DIR/docs/" "s3://$BUCKET/" \
  --delete \
  --cache-control "public,max-age=300" \
  --exclude ".nojekyll"

aws s3 cp "$ROOT_DIR/docs/index.html" "s3://$BUCKET/index.html" \
  --content-type "text/html; charset=utf-8" \
  --cache-control "public,max-age=300" \
  --metadata-directive REPLACE

aws cloudfront create-invalidation \
  --distribution-id "$DISTRIBUTION_ID" \
  --paths "/*" \
  --query "Invalidation.Id" \
  --output text
