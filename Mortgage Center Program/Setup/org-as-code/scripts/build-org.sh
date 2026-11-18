#!/usr/bin/env bash
# Build the Apex Home Lending org from source in one go.
#   scripts/build-org.sh <alias>             deploy to an org you already logged in to
#   scripts/build-org.sh <alias> --scratch   create a 7-day scratch org from the Dev Hub first
set -euo pipefail
ORG="${1:?usage: scripts/build-org.sh <alias> [--scratch]}"
DEV_HUB="${DEV_HUB:-apex-dev}"

if [[ "${2:-}" == "--scratch" ]]; then
  sf org create scratch --definition-file config/project-scratch-def.json \
    --alias "$ORG" --target-dev-hub "$DEV_HUB" --duration-days 7 --wait 15
fi

# One deploy: the Metadata API works out the order (objects before fields, fields before roll-ups and formulas)
sf project deploy start --target-org "$ORG" --manifest manifest/package.xml --wait 30

sf org assign permset --target-org "$ORG" \
  --name Apex_Data_Model_1 --name Apex_Data_Model_2 --name Apex_Data_Model_3 --name Apex_Data_Model_4

if [[ -f data/sample-plan.json ]]; then
  sf data import tree --target-org "$ORG" --plan data/sample-plan.json
fi

echo "Built $ORG. Opening it…"
sf org open --target-org "$ORG"
