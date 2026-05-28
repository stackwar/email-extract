# Email Extract

从 QQ 邮箱中自动提取骑手信息并整理到 Excel 表格。

## 功能

- 自动连接 QQ 邮箱，提取含"聊天记录"的邮件
- 支持文字邮件和图片邮件（OCR 识别）
- 按年月自动分 sheet，按姓名+电话去重
- 支持 IMAP IDLE 监听模式，新邮件到达自动处理
- 授权码验证，支持过期控制

## 使用

```bash
pip install -r requirements.txt

# 单次提取
python extract.py

# 监听模式（常驻后台，自动处理新邮件）
python extract.py watch
```

## 配置

复制 `config.example.json` 为 `config.json`，填入邮箱地址和授权码。

## 构建

推送版本 tag（如 `v1.0.0`）自动触发 GitHub Actions 构建 Windows EXE。
