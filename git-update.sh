#!/bin/bash
# git-update.sh - Automatically add, commit, and push all changes

# Get the current date and time for the commit message
TIMESTAMP=$(date +"%Y-%m-%d %H:%M:%S")

# Change to the repository directory (uncomment if needed)
# cd /Users/mpriessner/windsurf_repos/dictate_tool

# Add all changes
git add .

# Commit with timestamp
git commit -m "Updated version - $TIMESTAMP"

# Push to the current branch
CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
git push origin $CURRENT_BRANCH

echo "✅ Changes committed and pushed to $CURRENT_BRANCH"
