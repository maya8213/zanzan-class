# 赞赞轻松管班

班主任工作减负神器，极简操作 · 智能评分 · 一键AI生成各种工作记录文档。

## 功能模块

- 👥 学生名单管理
- 📋 考勤记录
- 🍽️ 就餐记录
- ⭐ 课堂表现
- 📚 课程管理
- 🎤 班会记录
- 💬 家长沟通记录
- 📈 学情反馈报告
- 📝 班务日志生成

## 快速开始

```bash
pip install flask openpyxl python-docx

python app.py
```

访问 http://localhost:5000

## 部署到学校内网

```bash
# 在服务器上运行（需要内网IP）
python app.py
# 手机/其他设备访问 http://<服务器IP>:5000
```

## 技术栈

- 后端：Flask + SQLite
- 前端：HTML + JavaScript（手机端无需安装任何APP）
- 文档生成：python-docx（Word文档）
