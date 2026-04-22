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

git push

rm -f "$STATE_FILE"

echo "DONE"
