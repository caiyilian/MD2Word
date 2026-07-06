# GitHub CLI 中文编码问题

## 问题

在 Windows 上使用 `gh` CLI 创建/编辑 Issue 或 PR 时，如果内容包含中文（或其它非 ASCII 字符），
`gh` 会以系统活动代码页（如 GBK/Windows-1252）而非 UTF-8 编码发送请求，
导致 GitHub 上的中文内容出现乱码（如 `使用` 变成 `使�?` 或 `使�`）。

## 根因

`gh` CLI 在 Windows 上依赖 PowerShell 的管道和参数传递。PowerShell 5.1 默认输出为 UTF-16LE，
但 `gh` 处理输入时使用系统代码页（中国大陆为 GBK 936），导致 UTF-8 编码的中文被二次编码损坏。

## 涉及命令

以下命令均可能触发此问题：

```bash
# 全部有风险
gh issue create --title "中文标题" --body "中文内容"
gh issue edit <num> --body "中文内容"
gh pr create --title "中文标题" --body "中文内容"
gh pr merge <num> --subject "中文标题"
```

## 安全操作方式

### 方式一：使用 gh api + --input（推荐）

```python
python -c "
import subprocess, json
payload = json.dumps({'title': '中文标题', 'body': '中文内容'}).encode('utf-8')
subprocess.run(['gh', 'api', 'repos/:owner/:repo/issues', '--method', 'POST', '--input', '-'],
               input=payload)
"
```

`json.dumps()` 输出纯 UTF-8 字节流，`--input -` 从 stdin 读取，不走系统代码页。

### 方式二：Python 的 subprocess + UTF-8 编码

```python
import subprocess
payload = json.dumps({'body': '中文'}).encode('utf-8')
subprocess.run(['gh', 'api', 'repos/:owner/:repo/issues/N', '--method', 'PATCH', '--input', '-'],
               input=payload)
```

### 方式三：使用 GitHub REST API 直接调用

```python
import requests
headers = {'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json'}
requests.patch(f'https://api.github.com/repos/owner/repo/issues/N',
               json={'title': '中文'}, headers=headers)
```

### 方式四：PowerShell 中先用文件（风险）

```powershell
Set-Content -Path body.md -Value "中文" -Encoding utf8
gh issue edit <num> --body-file body.md
```

注意：`gh` 在 Windows 上读取文件时可能仍用 GBK 编码，此方式不一定可靠。
