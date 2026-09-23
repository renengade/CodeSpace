#!/usr/bin/env bash
#
# getcatalogue 一键安装脚本（定制版）
# 用法：
#   curl -fsSL https://raw.githubusercontent.com/renengade/CodeSpace/main/install_getcatalogue.sh | bash
#   curl -fsSL .../install_getcatalogue.sh | bash -s -- --uninstall   # 卸载
#
set -euo pipefail

# ================== 配置区（已按仓库实际情况填写） ==================
REPO_RAW="${GET_CATALOGUE_RAW:-https://raw.githubusercontent.com/renengade/CodeSpace/main}"
APP_NAME="getcatalogue"         # 仓库中的文件名（无下划线）
CMD_NAME="getc"        # 用户在终端输入的命令名
PY_MIN=(3 9)                    # pandas 2.x 要求 Python >= 3.9
DEPS="pandas openpyxl"          # 仓库暂无 requirements.txt，用内置依赖列表
                                # openpyxl 是 df.to_excel() 的必需引擎
INSTALL_DIR="$HOME/.local/share/$APP_NAME"
BIN_DIR="$HOME/.local/bin"
VENV_DIR="$INSTALL_DIR/.venv"
# ====================================================================

info()  { printf '  \033[32m✓\033[0m %s\n' "$*"; }
warn()  { printf '  \033[33m!\033[0m %s\n' "$*"; }
die()   { printf '  \033[31m✗\033[0m %s\n' "$*" >&2; exit 1; }

# ---------- 卸载 ----------
if [[ "${1:-}" == "--uninstall" ]]; then
  rm -rf "$INSTALL_DIR" "$BIN_DIR/$CMD_NAME"
  info "已卸载 $APP_NAME"
  exit 0
fi

# ---------- 1. 检查 Python ----------
command -v python3 >/dev/null 2>&1 || die "未找到 python3，请先安装 Python ${PY_MIN[0]}.${PY_MIN[1]}+"
python3 -c "import sys; sys.exit(0 if sys.version_info >= (${PY_MIN[0]}, ${PY_MIN[1]}) else 1)" \
  || die "Python 版本过低，需要 ${PY_MIN[0]}.${PY_MIN[1]}+（当前：$(python3 -V 2>&1)）"
info "Python $(python3 -V 2>&1 | awk '{print $2}')"

# ---------- 2. 下载程序 ----------
mkdir -p "$INSTALL_DIR"
command -v curl >/dev/null 2>&1 || die "未找到 curl，请先安装"
info "下载 $APP_NAME.py ..."
curl -fsSL "$REPO_RAW/$APP_NAME.py" -o "$INSTALL_DIR/$APP_NAME.py"

# ---------- 3. 创建 venv 并安装依赖 ----------
if ! python3 -m venv "$VENV_DIR" 2>/dev/null; then
  die "创建虚拟环境失败。Debian/Ubuntu 请先执行：sudo apt install python3-venv python3-pip"
fi

PIP="$VENV_DIR/bin/pip"
"$PIP" install --upgrade pip -q

if curl -fsSL "$REPO_RAW/requirements.txt" -o "$INSTALL_DIR/requirements.txt" 2>/dev/null; then
  info "安装依赖（requirements.txt）..."
  "$PIP" install -q -r "$INSTALL_DIR/requirements.txt"
else
  warn "仓库暂无 requirements.txt，使用内置依赖列表：$DEPS"
  # shellcheck disable=SC2086
  "$PIP" install -q $DEPS
fi
info "依赖安装完成（pandas + openpyxl）"

# ---------- 4. 生成终端命令入口 ----------
# 关键：默认目录取用户当前目录（-d 未指定时），而不是脚本安装目录
mkdir -p "$BIN_DIR"
cat > "$BIN_DIR/$CMD_NAME" <<EOF
#!/usr/bin/env bash
# getcatalogue 启动器（由安装脚本生成）
# 未指定 -d 时，把当前目录作为默认目标目录
exec "$VENV_DIR/bin/python" "$INSTALL_DIR/$APP_NAME.py" "\$@"
EOF
chmod +x "$BIN_DIR/$CMD_NAME"
info "创建命令：$BIN_DIR/$CMD_NAME"

# ---------- 5. PATH 检查 ----------
if ! case ":$PATH:" in *":$BIN_DIR:"*) true ;; *) false ;; esac; then
  warn "$BIN_DIR 不在 PATH 中，请将下面这行加入 ~/.bashrc 或 ~/.zshrc："
  printf '      export PATH="$HOME/.local/bin:$PATH"\n\n'
fi

echo
info "getcatalogue 安装完成！打开新终端或执行 source ~/.bashrc 后使用："
cat <<'EOF'

      get-catalogue xlsx               # 列出当前目录下的 xlsx 清单并导出
      get-catalogue pdf -d /path/dir   # 指定目录和文件类型
      get-catalogue folder -n          # 仅打印子文件夹，不导出
      get-catalogue                    # 交互模式

EOF
