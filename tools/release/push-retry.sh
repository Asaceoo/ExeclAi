#!/bin/bash
cd /d/execl1.0-wps
LOG=/d/execl1.0-wps/_dev_tmp/push-retry.log
TARGET=$(git rev-parse github-main)
echo "=== retry loop start $(date) target=$TARGET ===" >> $LOG
for i in $(seq 1 60); do
  echo "--- attempt $i $(date +%H:%M:%S) ---" >> $LOG
  git push git@github.com:Asaceoo/ExeclAi.git github-main:main --force --progress >> $LOG 2>&1
  REMOTE=$(git ls-remote git@github.com:Asaceoo/ExeclAi.git main 2>/dev/null | cut -f1)
  if [ "$REMOTE" = "$TARGET" ]; then
    echo "PUSH SUCCESS attempt $i $(date)" >> $LOG
    exit 0
  fi
  sleep 600
done
echo "ALL 60 ATTEMPTS FAILED $(date)" >> $LOG
exit 1
