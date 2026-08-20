# NPA Meeting Minutes Skill

将中文会议转写、讨论记录或粗略笔记整理为新力通标准格式的 Word 会议纪要。

## 主要能力

- 提炼会议背景、讨论内容、共识决议和行动项
- 生成带公司页眉、页脚、会议信息区和签批区的 `.docx`
- 使用单一固定宽度主表，保证 WPS 中正文边框连续、左右边界对齐
- 提供结构校验脚本，检查主表网格、正文通栏和签批区边框

## 安装

将本仓库克隆到 Codex Skills 目录：

```powershell
git clone https://github.com/sxt916/npa-meeting-minutes-skill.git `
  "$env:USERPROFILE/.codex/skills/npa-meeting-minutes"
```

## 使用

在 Codex 中提供会议原始材料，并说明：

```text
调用 npa-meeting-minutes skill，整理成标准格式会议纪要。
```

详细输入格式、生成流程和验证要求见 `SKILL.md`。

## 目录

- `SKILL.md`：技能说明和工作流
- `agents/openai.yaml`：技能展示信息
- `assets/`：公司页眉、页脚素材
- `scripts/build_minutes_docx.py`：DOCX 生成器
- `scripts/check_border_continuity.py`：WPS 对齐和边框连续性校验器
