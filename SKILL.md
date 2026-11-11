---
name: office-proofing-cleaner
description: 统一处理 Word .docx 与 PowerPoint .pptx 的红色或蓝色校对波浪线，按文件格式选择方法，保留原稿及正文版式。适用于用户要求去除 Office 文字下方拼写或语法波浪线、合并处理 Word 与 PPT 文件；不用于删除文字自身的下划线或校正真实错别字。
---

# Office 红色波浪线处理

使用一个入口处理 Word 和 PPT，无需用户分别选择旧技能。先识别文件格式，再采用对应方法；统一入口不代表两个格式具有相同的校对机制。

## 方法与边界

| 文件 | 默认方法 | 后续拼写检查 |
| --- | --- | --- |
| `.pptx` | 修复文字校对语言，重置旧错误标记 | 保持启用；新错词仍应被识别 |
| `.docx` | 沿用原 Word 技能：清理错误标记，隐藏当前文档的校对提示，并设置当前文字免校对 | 当前文档/文字的校对受到抑制，不能声称仍正常检查 |

两种方法都不修改 Office 全局选项或用户词典。Word 模式会修改**文档内部**校对设置；若用户要求 Word 也保持正常校对、或不设置免校对，不运行该模式。先按其要求调查语言及状态问题，说明现有 Word 清理脚本不能满足该约束，不能以隐藏错误代替修复。

## 常用操作

所有相对脚本路径均相对于本技能目录。

```powershell
# 只检查，不改文件
python scripts/clean_office_proofing.py "D:\材料\文件.docx" --inspect

# 根据扩展名处理 Word 或 PPT，另存带编号的独立副本
python scripts/clean_office_proofing.py "D:\材料\文件.docx"
python scripts/clean_office_proofing.py "D:\材料\文件.pptx"

# 指定一个尚不存在的输出文件
python scripts/clean_office_proofing.py "D:\材料\原稿.pptx" "D:\材料\处理副本.pptx"
```

- 默认保留源文件，另存 `原名_校对处理副本`，重名自动编号；拒绝覆盖源文件、已有输出或报告。
- 用户指定文件夹时，仅处理获授权范围内的 `.docx`、`.pptx`；默认不递归，排除 `~$` 锁文件、处理副本和检查目录。先列定源文件，再逐个处理，避免把本次输出当成新输入。
- 以完整路径识别已打开的源文件。需要包含未保存内容时先建立独立快照，不保存或关闭原稿，不根据同名文件猜测。
- 签名、加密、带宏和旧版 `.doc`/`.ppt` 文件不直接交给脚本转换；保留源文件并报告限制。不改文字、不接受或拒绝修订、不移除评论、图片或媒体。

## 按格式加载与验证

处理 PPT 时读取 [references/ppt-workflow.md](references/ppt-workflow.md)。所需修复、PowerPoint 验证及测试脚本已包含在本技能中，无需调用旧技能。遇到语言、引号显示或重复红线时再读取 [references/ppt-proofing-notes.md](references/ppt-proofing-notes.md)。

处理 Word 时读取 [references/word-workflow.md](references/word-workflow.md)。Word 脚本需要本机已有的 `lxml`；PPT 基础修复只用 Python 标准库，不因处理 PPT 而安装 Word 依赖。

统一脚本会校验正文、非校对格式、非相关部件及源文件哈希，并生成旁边的 `.validation.json`。这只证明文件处理结果，不能代替 Word/PowerPoint 应用内验证。能使用本机 Office 时打开独立副本检查；用户正在编辑 Office 时不要抢占窗口或丢弃测试窗口中的人工输入，无法完成的检查如实标注。

交付时给出独立副本链接，说明文件格式、采用的方法、原稿保护情况、已完成的验证及剩余问题。Word 结果明确写“当前文档免校对”；PPT 结果只有在重新检查和新错词测试通过后，才可说“保留正常校对且红线已修复”。
