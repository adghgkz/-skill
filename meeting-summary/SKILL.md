---
name: meeting-summary
description: This skill should be used when the user requests generating meeting summaries, meeting minutes, or meeting records in Word format.
---

# 会议总结技能

本技能用于生成结构化的会议总结 Word 文档。

## 使用场景

- 例会、周会、月会总结
- 项目评审会议记录
- 团队内部讨论会议
- 任何需要生成会议纪要的场景

## 触发关键词

- "会议总结"、"会议记录"、"会议纪要"
- "meeting summary"、"meeting minutes"
- "生成会议文档"

## 输入信息

用户需要提供以下信息：

| 字段 | 说明 | 示例 |
|------|------|------|
| meeting_date | 会议日期 | 2026-03-23 |
| meeting_time | 会议时间 | 14:00-15:30 |
| attendees | 参会人员 | 张三、李四、王五 |
| agenda | 会议议题 | 项目进度评审、下季度计划 |
| content | 讨论内容 | 各项目进展、问题讨论 |
| conclusions | 会议结论 | 同意进入下一阶段 |
| action_items | 行动项 | 任务+负责人+截止日期 |

## 输出格式

生成 Word 文档 (.docx)，包含：

1. **会议基本信息**
   - 会议标题
   - 日期、时间
   - 地点
   - 参会人员

2. **会议背景**
   - 会议目的
   - 讨论背景

3. **讨论内容**
   - 各议题讨论要点

4. **会议结论**
   - 决策事项
   - 共识达成

5. **行动项**
   - 任务描述
   - 负责人
   - 截止日期

## 使用方法

当用户请求生成会议总结时：

1. 收集完整的会议信息
2. 识别会议类型（例会/周会/月会/评审会）
3. 调用 `scripts/generate_summary.py` 生成 Word 文档
4. 将生成的文档路径告知用户

## 脚本调用

```bash
python scripts/generate_summary.py \
    --date "2026-03-23" \
    --time "14:00-15:30" \
    --location "会议室A" \
    --attendees "张三,李四,王五" \
    --agenda "项目进度评审" \
    --content "讨论内容..." \
    --conclusions "结论..." \
    --actions "任务1|张三|2026-03-30,任务2|李四|2026-04-05" \
    --output "会议总结_2026-03-23.docx"
```

## 依赖

- python-docx 库
- 本技能依赖 `docx` skill 进行 Word 文档生成
