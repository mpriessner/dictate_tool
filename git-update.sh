#!/bin/bash
# git-update.sh - Automatically add, commit, and push all changes

# Get the current date and time for the commit message
TIMESTAMP=$(date +"%Y-%m-%d %H:%M:%S")

# Check if a commit message was provided as an argument
if [ $# -eq 0 ]; then
  # No argument provided, use the default timestamp message
  COMMIT_MSG="Updated version - $TIMESTAMP"
else
  # Use the provided message with timestamp
  COMMIT_MSG="$1 - $TIMESTAMP"
fi

# Add all changes
git add .

# Commit with the message
git commit -m "$COMMIT_MSG"

# Push to the current branch
CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
git push origin $CURRENT_BRANCH

echo "✅ Changes committed and pushed to $CURRENT_BRANCH with message: $COMMIT_MSG"