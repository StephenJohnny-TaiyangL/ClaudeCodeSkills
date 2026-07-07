# Convert MathType Skill

一键将Word文档中的MathType公式批量转换为Word原生可编辑公式(OMML格式)。

## 功能说明

将 `.docx` 文件中所有 MathType (Equation.DSMT4) OLE嵌入公式，通过 MTEF→MathML→OMML 管道自动转换为Word原生可编辑公式，支持分数、上下标、矩阵、根号等所有复杂公式结构。

## 安装方法

1. 将 `convert-mathtype-skill` 整个文件夹复制到你的 skill 目录：
   - Kiro: `C:\Users\你的用户名\.kiro\skills\`
   - Claude Code: 项目根目录下 `.claude\skills\`

2. 将 `convert_mathtype.py` 脚本放到任意位置（建议放在固定目录如 `D:\PythonProjects\`）

3. 修改 `convert-mathtype.md` 中脚本路径为你的实际路径

## 前置条件

- Windows 系统
- MathType 已安装（需要其 COM 接口）
- Microsoft Office 已安装（需要 MML2OMML.XSL）
- Python 3.8+ 并安装以下包：
  ```
  pip install pywin32 olefile lxml
  ```

## 使用方式

在对话中输入：
```
/convert-mathtype C:\path\to\your\file.docx
```

或直接命令行运行：
```
python convert_mathtype.py "C:\path\to\your\file.docx"
```

## 许可

MIT License - 可自由使用、修改和分发。
