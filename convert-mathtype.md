---
name: convert-mathtype
description: 将Word文档(.docx)中的MathType公式批量转换为Word原生可编辑OMML格式
---

# Convert MathType to OMML

将指定Word文档中所有MathType (Equation.DSMT4) OLE公式转换为Word原生可编辑公式(OMML格式)。

## 使用方式

用户提供 `.docx` 文件的完整路径作为参数。如果未提供路径，请询问用户。

## 执行步骤

1. 确认文件存在且为 `.docx` 格式
2. 修改 `d:\PythonProjects\convert_mathtype.py` 脚本中的 `INPUT_FILE` 和 `OUTPUT_FILE` 为用户提供的路径
3. 运行脚本：`cd d:/PythonProjects && python convert_mathtype.py`
4. 使用 Word COM 自动化验证转换结果：
   - 确认 MathType OLE 公式数量为 0
   - 确认 OMML 公式数量正确
   - 确认文档能正常打开
5. 向用户报告：
   - 转换的公式数量
   - 验证状态
   - 备份文件位置（原文件名 + `.bak`）

## 前置条件

- 系统已安装 MathType（用于 COM 接口 `Equation.DSMT4`）
- 系统已安装 Microsoft Office（用于 `MML2OMML.XSL` 和验证）
- Python 环境已安装：`win32com`, `olefile`, `lxml`
- 转换脚本位于 `d:\PythonProjects\convert_mathtype.py`

## 注意事项

- 脚本会自动创建 `.bak` 备份文件（仅首次转换时）
- 如果转换失败数量 > 0，需要检查具体原因并告知用户
- 转换管道：MTEF二进制 → MathML（MathType COM）→ OMML（XSL转换）
