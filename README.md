# Office 红色波浪线处理

**一个 Codex 技能，处理 Word 和 PowerPoint 的校对波浪线。** 遇到中文 PPT 大片红线、复制粘贴后校对语言不对，或只想交付一份不显示校对提示的 Word 副本，可以把文件交给 `$office-proofing-cleaner`。它按格式选择方法，保留原稿，另存一份可编辑的文件和校验报告。

## 两种格式，两种处理方式

| 格式 | 怎么处理 | 后续校对 |
| --- | --- | --- |
| PowerPoint `.pptx` | 修复可识别的校对语言问题，重置旧的拼写错误标记 | 保留校对功能；真实错词、网址和专业术语仍可能被标记 |
| Word `.docx` | 清理校对错误标记，隐藏当前文档的红蓝提示，并设置现有文字免校对 | 当前文档的校对受到抑制；如果要保留 Word 文档正常校对，请不要用此模式 |

两种方式都不修改 Office 全局拼写选项或个人词典。文件在本机处理，无需上传文档。脚本会核对文字、非校对格式及未涉及的文件部件，默认不覆盖源文件；同名输出会自动编号。

## 三步开始

1. 在本仓库点击 **Code → Download ZIP**，解压后将目录命名为 `office-proofing-cleaner`。
2. 将完整目录放到 `用户主目录/.agents/skills/`，使 `office-proofing-cleaner/SKILL.md` 与 `scripts/`、`references/`、`agents/` 同级。可参考 [Codex 官方技能目录说明](https://learn.chatgpt.com/docs/build-skills#where-to-save-skills)。
3. 在 Codex 里附上文档或提供其本机路径，然后发送：

   > 使用 $office-proofing-cleaner 处理这份 PPT 的红色波浪线，保留正常拼写检查，另存副本并告诉我验证结果。

   处理混合文件夹时也可说：“使用 $office-proofing-cleaner 处理这个文件夹里的 Word 和 PPT，保留原稿。”

如果技能没有出现在列表中，先确认没有多套一层目录，再重启 Codex。

## 不用 Codex，也可以运行

建议使用 Python 3.10 或以上版本。PPT 基础处理只依赖标准库；处理 Word 时，在所用 Python 环境安装 `lxml`。

```text
python -m pip install lxml
```

在仓库目录中运行：

```text
python scripts/clean_office_proofing.py "演示文稿.pptx" --inspect
python scripts/clean_office_proofing.py "演示文稿.pptx"
python scripts/clean_office_proofing.py "报告.docx"
```

例如，`报告.docx` 会生成 `报告_校对处理副本.docx` 和对应的 `.validation.json`。输入文件也可以使用绝对路径。

## 如何判断处理成功

- **文件级检查**：报告列出处理部件、原稿是否未变，以及处理前后的校对状态。
- **Office 应用内检查**：导出的图片或 PDF 不显示校对红线，不能单独证明红线已消失。PowerPoint 应用内验证需要 Windows、桌面 PowerPoint、`pywin32` 和 `Pillow`；[验证指南](references/ppt-workflow.md)说明如何重开副本、对比逐页渲染并检查新错词仍能标记。Word 应用内可重开独立副本确认文档级设置生效。
- **支持范围**：只处理常规 `.docx` 和 `.pptx`。不处理旧版 `.doc`/`.ppt`、带宏、签名或加密文件，也不会改写真实错别字或删除文字自身的下划线。

运行附带的文件级测试：

```text
python -m unittest discover -s scripts -p "test_*.py"
```

## English overview

A local Codex skill and Python tool for proofing squiggles in Word and PowerPoint. It creates an editable copy and a validation report. **PowerPoint retains spell-checking; Word suppresses proofing in the processed document.** See [SKILL.md](SKILL.md) for the workflow and [the PowerPoint guide](references/ppt-workflow.md) for in-app verification.

The Word helper is adapted from [word-proofing-cleaner](https://github.com/czd176224-rgb/word-proofing-cleaner). Its original [MIT notice](licenses/word-proofing-cleaner-MIT.txt) is included for that component.
