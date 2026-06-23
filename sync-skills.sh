#!/bin/bash
# ============================================================
# Skills 自动同步脚本（直接模式）
# 位置：~/.claude/skills/ 本身就是 git 仓库
# 功能：检测变更 → 自动提交 → 推送到 GitHub tracy-skills
# ============================================================

SKILLS_DIR="$HOME/.claude/skills"
LOG_FILE="$SKILLS_DIR/sync-log.txt"
GH_PATH="/c/Program Files/GitHub CLI"

export PATH="$PATH:$GH_PATH"

echo "==== $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG_FILE"

cd "$SKILLS_DIR" || exit 1

# 清理意外带入的 .git 子目录
find . -type d -name ".git" -not -path "./.git" -exec rm -rf {} + 2>/dev/null

# 检测变更
git add -A >> "$LOG_FILE" 2>&1

if git diff --staged --quiet; then
    echo "无变更" >> "$LOG_FILE"
    exit 0
fi

# 提交并推送
git commit -m "自动同步: $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
git push origin master >> "$LOG_FILE" 2>&1

echo "同步完成 ✓" >> "$LOG_FILE"
