#!/usr/bin/env bash
# Register the open-skills Office skills in the LiteLLM Skill Hub, then publish
# them to the public hub.
#
# Requires a proxy ADMIN key. A scoped virtual key (role internal_user_viewer)
# is rejected with 401 on /claude-code/plugins.
#
#   LITELLM_ADMIN_KEY=sk-... ./register-skill-hub.sh
#
# Consumers then add the marketplace once, in ~/.claude/settings.json:
#
#   "extraKnownMarketplaces": {
#     "volpe": { "source": "url", "url": "http://10.75.42.137:4000/claude-code/marketplace.json" }
#   }
#
# ...and install with:  /plugin marketplace add docx-creation-editing

set -euo pipefail

PROXY="${LITELLM_BASE_URL:-http://10.75.42.137:4000}"
KEY="${LITELLM_ADMIN_KEY:?set LITELLM_ADMIN_KEY to a proxy admin key}"
REPO="https://github.com/shauneshraghi/open-skills"

register() {
  local name=$1 desc=$2
  echo "→ $name"
  curl -fsS -X POST "$PROXY/claude-code/plugins" \
    -H "Authorization: Bearer $KEY" \
    -H "Content-Type: application/json" \
    -d @- <<JSON
{
  "name": "$name",
  "source": { "source": "git-subdir", "url": "$REPO", "path": "$name" },
  "description": "$desc",
  "domain": "Documents",
  "namespace": "office"
}
JSON
  echo
  curl -fsS -X POST "$PROXY/claude-code/plugins/$name/enable" \
    -H "Authorization: Bearer $KEY"
  echo
}

register docx-creation-editing \
  "Create and edit Word .docx files via python-docx and raw OOXML, with the Word VBA reference bundled"
register pptx-creation-editing \
  "Create and edit PowerPoint .pptx files via python-pptx and raw OOXML, with the PowerPoint VBA reference bundled"
register xlsx-creation-editing \
  "Create and edit Excel .xlsx files via openpyxl and raw OOXML, with the Excel VBA reference bundled"

echo "=== public hub ==="
curl -fsS "$PROXY/public/skill_hub"
echo
echo "=== marketplace manifest ==="
curl -fsS "$PROXY/claude-code/marketplace.json"
