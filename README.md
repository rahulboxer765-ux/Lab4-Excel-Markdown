# Lab 4: Working with Excel and Markdown

## Introduction

This article describes my experience using PowerShell and the ImportExcel module on a Windows Server virtual machine in Hyper-V.

### Learning Objectives

1. Import data from Excel.
2. Export process information to Excel.
3. Format worksheets and create charts.
4. Document the lab using Markdown and Git.

## Markdown Text Formatting

**Bold text** highlights important information.

*Italic text* emphasizes a word or phrase.

***Bold and italic*** combines both styles.

~~Strikethrough~~ marks outdated information.

## Tools Used

- Windows Server
- Hyper-V
- PowerShell
- ImportExcel module
- Git Bash
- GitHub

## Task Checklist

- [x] Install ImportExcel
- [x] Create Fruits.xlsx
- [x] Import Excel data
- [x] Export running processes
- [x] Create a fruit sales chart
- [ ] Publish this article to GitHub

## PowerShell Commands

The `Import-Excel` command reads spreadsheet data.

```powershell
Install-Module -Name ImportExcel
Import-Excel "C:\Lab4\Fruits.xlsx"
Get-Process | Export-Excel "C:\Lab4\ProcessInfo.xlsx" -WorksheetName "Processes"
```

## Fruit Sales Results

| Fruit | Total Sold |
|:------|-----------:|
| Apple | 50 |
| Banana | 35 |
| Orange | 40 |
| Mango | 25 |

## Important Observation

> The ImportExcel module can create and modify Excel workbooks without Microsoft Excel installed on Windows Server.

## Links

Learn more about [GitHub](https://github.com) and [Markdown syntax](https://www.markdownguide.org/basic-syntax/).

## Image Syntax

![Markdown logo](https://upload.wikimedia.org/wikipedia/commons/4/48/Markdown-mark.svg)

## Horizontal Rule

---

## Conclusion

This lab helped me practice importing, exporting, and formatting Excel data using PowerShell. I also learned how to document technical work using Markdown.
