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
# Enforce Conventional Commit style

commit_msg_file=$1
commit_msg=$(cat "$commit_msg_file")

# Conventional Commits: type(scope?): description
regex="^(feat|fix|docs|style|refactor|perf|test|chore)(\([a-z0-9_-]+\))?: .+"

if ! echo "$commit_msg" | grep -qE "$regex"; then
  echo "❌ Commit message does not follow Conventional Commits."
  echo "   Format: <type>(<scope>): <description>"
  echo "   Example: feat(auth): add login validation"
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