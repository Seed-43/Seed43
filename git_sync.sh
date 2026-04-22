#!/bin/bash

set -x

echo "=== SCRIPT STARTED ==="

if [ ! -d ".git" ]; then
  echo "NOT A GIT REPO (run inside Seed43 folder)"
  exit 1
fi

STATE_FILE=".git_sync_state"

echo "Current folder:"
pwd

# 🔐 TOKEN (you fill this in)
GITHUB_TOKEN=SHA256:8NMuP/DuNTPXr8Kk/KXQWNdzHCMyO/LHrme+l1aTu0A

if [ ! -f "$STATE_FILE" ]; then
  echo "FIRST RUN → PULL MODE"

  git pull

  echo "READY" > "$STATE_FILE"

  echo "DONE PULLING - now edit files"
  exit 0
fi

echo "SECOND RUN → PUSH MODE"

git status

git add -A

echo "Enter commit message:"
read msg

git commit -m "$msg"

# 🔐 Safe push using token
if [ -z "$GITHUB_TOKEN" ]; then
  echo "ERROR: GITHUB_TOKEN is not set"
  exit 1
fi

git push "https://Seed43:${GITHUB_TOKEN}@github.com/Seed-43/Seed43.git"

if [ $? -eq 0 ]; then
  rm -f "$STATE_FILE"
  echo "DONE"
else
  echo "PUSH FAILED - state file preserved"
  exit 1
fi
