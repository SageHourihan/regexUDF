# VBA RegX Function

The VBA `RegX` function is a custom function for Microsoft Excel that allows you to apply regular expressions to cell values and extract matching text. This function can be useful for tasks that involve searching for patterns within text data in Excel.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)
- [Escaping Double Quotes](#how-to-escape-double-quotes)

## Installation

### Installation within a Workbook

To use the `RegX` function in a specific Excel workbook, follow these steps:

1. Press `ALT + F11` to open the VBA editor in Excel.
2. Go to `Insert` -> `Module` to insert a new module.
3. Copy and paste the `RegX` function code into the module.
4. Save your Excel workbook.

The `RegX` function is now available for use within the workbook where it was defined.

### Installation Globally

To make the `RegX` function available globally in all your Excel workbooks, follow these steps:

1. Press `ALT + F11` to open the VBA editor in Excel and paste in the code.
2. Select `Tools` -> `References`. Check Microsoft `VBScript Regular Expressions 1.0` and `Microsoft VBScript Regular Expressions 5.5`
3. Save as `Excel Add In` to: C:\Users\userName\AppData\Roaming\Microsoft\AddIns
4. Close the Excel workbook.
5. Open Excel and go to `File` -> `Options`.
6. In the Excel Options window, select `Add-Ins` on the left sidebar.
7. In the "Add-Ins" section, choose "Excel Add-ins" from the drop-down menu and click the "Go..." button.
8. Click the "Browse..." button in the "Add-Ins" window and locate the VBA project file you exported earlier.
9. Select the VBA project file and click "OK."

The `RegX` function is now available globally in all your Excel workbooks.

## Usage

The `RegX` function takes two arguments:

1. `strInput` (String): The cell value or text string you want to search for a regular expression match.
2. `regexPattern` (Variant): The regular expression pattern you want to apply to `strInput`.

The function returns the first match found in `strInput` based on the `regexPattern`.

In the example below, the function is applied to cell A1, searching for the pattern \d{3}-\d{2}-\d{4} (a common format for Social Security Numbers). If a match is found, it returns the matched text; otherwise, it returns "not matched."

```excel
=RegX(A1, "\d{3}-\d{2}-\d{4}")
```
## Examples

Here are some examples on how to use the `RegX` function 

1. Extracting dates
```excel
=RegX(A1, "\d{2}/\d{2}/\d{4}")
```

2. Finding Email Addresses
```excel
=RegX(A1, "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}")
```

## How To Escape Double Quotes
To include double quotes within your regular expression pattern, you should escape them by doubling them. For example, to match the text "abc", your regular expression pattern would look like this: "abc"""".
```excel
=RegX(A1, """[^""]+""")
```
