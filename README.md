# 📚 Scholara – AI Literature Review Platform

**Live URL:** [https://scholara.streamlit.app](https://scholara.streamlit.app) _(once deployed)_

**Author:** Soban | NUST MS Artificial Intelligence | July 2026

---

## 🎯 Project Overview

Scholara is a research‑focused web application that helps academics and researchers generate structured literature reviews from uploaded PDFs. It uses **Claude AI** (Anthropic) via the API to read, synthesise, and write a comprehensive review with introduction, thematic sections, contradiction detection, research gaps, and conclusion.

**Core Promise:** Upload papers → AI reads and synthesises → Download a citation‑ready review.

---

## ✨ Features

| Feature                 | Description                                                         |
| ----------------------- | ------------------------------------------------------------------- |
| **Multi‑PDF Upload**    | Upload up to 50 PDFs at once                                        |
| **AI Synthesis**        | Claude API extracts themes, detects contradictions, identifies gaps |
| **Structured Output**   | Introduction, thematic sections, gap analysis, conclusion           |
| **Citation Styles**     | APA 7, IEEE, MLA 9, Chicago (configurable)                          |
| **Export Options**      | Download as `.txt` or `.docx` (Word)                                |
| **Review History**      | Session‑based history of all generated reviews                      |
| **No Sign‑up Required** | Try the demo immediately                                            |

---

## 🧠 System Prompt (Used in API Call)

The following prompt is sent to Claude to ensure academic quality:
