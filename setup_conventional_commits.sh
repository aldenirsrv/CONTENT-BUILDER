#!/bin/bash
# =========================================================
# Setup Git commit-msg hook to enforce Conventional Commits
# =========================================================

set -e

HOOK_DIR=".git/hooks"
HOOK_FILE="$HOOK_DIR/commit-msg"

echo "🚀 Setting up Conventional Commit hook..."

# Ensure we are inside a Git repo
if [ ! -d ".git" ]; then
  echo "❌ This is not a Git repository (no .git directory found)."
  exit 1
fi

# Create hook directory if missing
mkdir -p "$HOOK_DIR"

# Write the commit-msg hook
cat > "$HOOK_FILE" <<'HOOK'
#!/bin/sh
# =========================================================
# Enforce Conventional Commit style
#
# Format: <type>(<scope>): <description>
#
# Types:
#   feat     → A new feature
#   fix      → A bug fix
#   docs     → Documentation only changes
#   style    → Changes that do not affect code meaning
#              (formatting, missing semi-colons, etc.)
#   refactor → Code change that neither fixes a bug nor adds a feature
#   perf     → Performance improvements
#   test     → Adding or correcting tests
#   chore    → Maintenance tasks (build, tooling, configs, etc.)
# =========================================================

commit_msg_file=$1
commit_msg=$(cat "$commit_msg_file")

# Conventional Commits regex
regex="^(feat|fix|docs|style|refactor|perf|test|chore)(\([a-z0-9_-]+\))?: .+"

if ! echo "$commit_msg" | grep -qE "$regex"; then
  echo "❌ Commit message does not follow Conventional Commits."
  echo
  echo "   Format: <type>(<scope>): <description>"
  echo
  echo "   Examples:"
  echo "     feat(auth): add login validation"
  echo "     fix(api): handle null values in response"
  echo "     docs(readme): update setup instructions"
  echo "     style(css): fix button alignment"
  echo "     refactor(core): simplify data pipeline"
  echo "     perf(query): optimize DB lookup speed"
  echo "     test(orders): add integration tests"
  echo "     chore(deps): update dependencies"
  echo
  exit 1
fi
HOOK

# Make it executable
chmod +x "$HOOK_FILE"

echo "✅ Conventional Commit hook installed at $HOOK_FILE"
echo
echo "Try it now:"
echo "   git commit -m \"feat(api): add new endpoint\"   # ✅ works"
echo "   git commit -m \"added new endpoint\"            # ❌ blocked"