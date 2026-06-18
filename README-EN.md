## Chaoxing Work Parser — Extract & Export to Word Exam Paper

No more manual copy-paste. No extra Python scripts. Extract questions from Chaoxing (Xuexitong) assignment pages in one click, and export directly as a **Word exam paper** (.docx) — ready for printing and final exam review.

Supported formats: **Word Exam Paper** / TXT / Markdown (ScriptCat / Tampermonkey).


### Installation
1. Install a userscript manager ([ScriptCat](https://docs.scriptcat.org/) or Tampermonkey)
2. Open the manager panel, click **"New"** or **"Add Script"**
3. Paste the contents of `学习通题目提取器.user.js` and save


### Usage (entirely in-browser)
1. Open a Chaoxing assignment page (make sure questions are loaded)
2. Click the **「提取题目」** floating button on the right side of the page
3. Click **「提取本页题目」** to extract all questions
4. Choose your output format:
   - **Word 试卷** — exports a formatted .docx exam paper, ready to print
   - **TXT** — plain text, with optional answers & wrong-question collection
   - **MD** — Markdown, with optional answers & wrong-question collection
5. Click **「下载」** (Download) or **「复制文本」** (Copy)


### Format Comparison

| Format | Use Case | Answers/Wrong Qs |
|--------|----------|------------------|
| Word Exam Paper | Print for review | Not supported (clean paper) |
| TXT | Archive, answer checking | Optional |
| Markdown | Import into note apps | Optional |


### How It Works
- Identifies question containers via DOM selectors `mark_item` / `questionLi`
- Extracts stems, options, and correct answers automatically
- Word documents are generated entirely in the browser using the docx.js library
- Output format matches manual copy-paste style