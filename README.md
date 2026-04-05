# FormatCodeBlocks

**LibreOffice Basic Macro**

This macro automates bulk code styling **TO SET THE DOCUMENT READY FOR HIGHLIGHTING** via the **Code Highlighter 2** extension in LibreOffice.

Once all the styles are applied to their respective blocks of code, selecting ```Highlight all codes formatted with the specified **paragraph style**,``` the extension can apply syntax highlighting to all designated text blocks **globally** in the document.

**User Guide**
Version 1.0 · File: `FormatCodeBlocks.bas`

---

## 1. Overview
> **Notes**
> 1. **Gemini** includes by default a previous line with the code’s name.
> 2. **ChatGpt** shows a not selectable note like "_</> Language name_" in the codebloks. Order ChatGpt to include a line with the code name before each block of code.
> 3. **Claude**: Order it to change the language identifier in the fenced code block syntax from lower case to capitalized names.

**FormatCodeBlocks** is a LibreOffice Writer macro that automates the formatting of code blocks across an entire document. It scans the document for specially named heading paragraphs (e.g. `Python`, `SQL`), selects all consecutive **Preformatted Text** paragraphs that follow, clears their direct formatting, and applies the matching `dk_Code_*` paragraph style.

Key capabilities:

- Processes the entire document in a single run
- Supports multiple languages in one execution
- Handles multiple code blocks of the same language
- Reports a summary count of paragraphs and blocks formatted
- Easily extensible: add new languages by editing one function
- Sets the document ready to apply code higlighting to all the document at once. (Tested using the _**Code Highlighter 2 - v2.7.4.4**_ extension)

---

## 2. Installation

1. Open LibreOffice Writer.
2. Go to **Tools > Macros > Organize Basic Dialogs**.
3. Select the library where you want to store the macro (e.g. `My Macros > Standard`) and click **Edit**.
4. In the Basic IDE, go to **File > Import Basic Source** and select `FormatCodeBlocks.bas` — or paste the entire file contents into a new module.
5. Save with **Ctrl+S**. The macro is now available.

> **Note:** The macro must be stored in a library that is loaded when Writer starts, such as `My Macros & Dialogs > Standard`.

---

## 3. Document Preparation

The macro looks for **exact text matches** — the heading paragraph must contain only the language name and nothing else. Follow this structure in your document:

### Required document structure

```
Python                          ← paragraph style: any
def hello():                    ← paragraph style: Preformatted Text
    print("Hello, world!")      ← paragraph style: Preformatted Text

SQL                             ← paragraph style: any
SELECT id, name FROM users      ← paragraph style: Preformatted Text
WHERE active = 1;               ← paragraph style: Preformatted Text
```

The paragraph containing `Python` or `SQL` can use any paragraph style — it is matched by its text content alone. The code lines below it must have the **Preformatted Text** paragraph style applied.

Rules:

- The language name must be the only text in its paragraph (leading/trailing spaces are ignored).
- The code paragraphs immediately below must use the **Preformatted Text** style.
- The block ends at the first paragraph that is not Preformatted Text, or at the next language heading, or at the end of the document.
- Multiple blocks of the same language are all processed.

---

## 4. Running the Macro

1. Open the document you want to process.
2. Go to **Tools > Macros > Run Basic Macro**.
3. Select `FormatCodeBlocks` from the macro list and click **Run**.
4. The language selection dialog will appear.
5. Select one or more languages (hold **Ctrl** to select multiple) and click **OK**.
6. The macro processes the document and shows a summary when finished.

> **Note:** You can assign the macro to a toolbar button or keyboard shortcut via **Tools > Customize** for faster access.

---

## 5. Language Selection Dialog

When the macro runs, a dialog displays a list of all available languages. You can:

- Click a single language to select it.
- Hold **Ctrl** and click to select multiple languages simultaneously _(The selection of not consecutive languages in the list is not possible currently)_.
- Click **OK** to process all selected languages, or **Cancel** to abort without changes.

If you click OK without selecting any language, the macro shows a warning and does not modify the document.

---

## 6. Supported Languages

The following languages are configured by default:

| Dialog Label | Heading Text in Document | Paragraph Style Applied |
|---|---|---|
| Bash | Bash | `dk_Code_Bash` |
| C++ | C++ | `dk_Code_Cpp` |
| CSS | CSS | `dk_Code_CSS` |
| Dart | Dart | `dk_Code_Dart` |
| HTML | HTML | `dk_Code_HTML` |
| Java | Java | `dk_Code_Java` |
| JavaScript | JavaScript | `dk_Code_JavaScript` |
| Plaintext | Plaintext | `dk_Code_Plaintext` |
| Python | Python | `dk_Code_Python` |
| SQL | SQL | `dk_Code_SQL` |
---

## 7. What the Macro Does

For each selected language, the macro performs the following steps:

### Step 1 — Scan

Enumerates every paragraph in the document looking for a paragraph whose text exactly matches the language name (e.g. `Python`).

### Step 2 — Collect

After finding the language heading, collects all consecutive paragraphs with the **Preformatted Text** style into a block.

### Step 3 — Clear direct formatting

For each paragraph in the block, a text cursor selects the full paragraph content and resets the following character properties to their style defaults:

- `CharWeight` (bold)
- `CharColor` (font colour)
- `CharHeight` (font size)
- `CharPosture` (italic)
- `CharUnderline`
- `CharBackColor` and `CharBackTransparent` (highlight)

The UNO dispatch command `ResetAttributes` is also applied to the selection for a thorough reset.

### Step 4 — Apply style

Sets the paragraph style of each collected paragraph to the corresponding `dk_Code_*` style (e.g. `dk_Code_Python`).

### Step 5 — Repeat

Continues scanning for the next occurrence of the same language heading and repeats steps 2–4. Then moves to the next selected language.

---

## 8. Adding New Languages

Open `FormatCodeBlocks.bas` in the Basic IDE and locate the `GetLanguages()` function at the top of the file:

```basic
Function GetLanguages() As Variant
    GetLanguages = Array( _
        Array("Bash",   "Bash",   "dk_Code_Bash"),   _
        Array("Python", "Python", "dk_Code_Python"), _
        Array("SQL",    "SQL",    "dk_Code_SQL")     _
    )
End Function
```

Each entry follows this pattern:

```basic
Array("Label in dialog", "Exact heading text in doc", "Paragraph style to apply")
```

To add Kotlin, for example, append a line:

```basic
Array("Kotlin", "Kotlin", "dk_Code_Kotlin")  _
```

**Important:** make sure the `dk_Code_Kotlin` paragraph style already exists in your document or document template before running the macro, otherwise LibreOffice will throw an error when trying to apply it.

---

## 9. Usage Examples

### Example A — Single language

Document contains one Python block:

```
Python                          ← paragraph style: any
import os                       ← paragraph style: Preformatted Text
print(os.getcwd())              ← paragraph style: Preformatted Text

Some normal text here.          ← block ends here
```

Run the macro, select **Python**, click OK. Result: both code lines receive `dk_Code_Python`. Summary: `"Formatted 2 paragraph(s) across 1 block(s)."`

### Example B — Multiple blocks, same language

Document contains two separate Python blocks:

```
Python
x = 1

Some text in between.

Python
y = 2
z = 3
```

Select **Python** and click OK. Result: all three code lines receive `dk_Code_Python`. Summary: `"Formatted 3 paragraph(s) across 2 block(s)."`

### Example C — Multiple languages in one run

Document contains both Python and SQL blocks. Select both in the dialog (**Ctrl+click**). The macro processes Python first, then SQL, and reports the combined total.

> _*Note*: Currently is not possible to select no consecutive languages in the list._

---

## 10. Troubleshooting

### "No matching blocks found"

The macro could not find a paragraph whose text exactly matches the selected language name. Check that the heading paragraph contains **only** the language name with no extra characters, and that the paragraphs below it use the **Preformatted Text** style.

### "Property or method not found" error

This usually means the macro is being run against a non-Writer document (e.g. Calc or Impress). The macro is designed for Writer only.

### Style not applied / error on style name

The `dk_Code_*` style does not exist in the current document. Create or import the style via **Format > Styles > Manage Styles** before running the macro.

### Only some paragraphs are styled

A paragraph inside the code block may have had its style manually changed away from **Preformatted Text** before running the macro. The macro stops collecting at the first paragraph that is not **Preformatted Text**. Restore the style on that paragraph and run the macro again.

---

*FormatCodeBlocks.bas · LibreOffice Basic Macro · v1.0*
