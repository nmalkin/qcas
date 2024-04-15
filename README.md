Qualitative Coding Assistant for Sheets
=======================================

QCAS (pronounced [quokkas](https://en.wikipedia.org/wiki/Quokka)) is the Qualitative Coding Assistant for Google Sheets.

## What problems does this solve?

[Qualitative coding](https://www.cessda.eu/Training/Training-Resources/Library/Data-Management-Expert-Guide/3.-Process/Qualitative-coding) is a process often used in [qualitative research](https://en.wikipedia.org/wiki/Qualitative_research), especially with the [grounded theory](https://en.wikipedia.org/wiki/Grounded_theory) methodology.
Basically, it involves a researcher (the coder or rater) looking at individual data items (survey responses, passages, images, etc.) and assigning them labels or categories based on their theme (or another dimension relevant for analysis). To put it even more simply, you look at a bunch of responses and assign labels to each.

Specialized software exists to aid with some variants of this process, but a quick and dirty way this analysis is often done is by having the responses in the column of a spreadsheet, then putting codes in the adjacent column.

This project is a set of [Google Apps Scripts](https://developers.google.com/apps-script/) that can be added to a Google Sheets spreadsheet to save time when performing these tasks.


## Features

**Type numbers instead of codes.** Instead of typing out the name of the code, you can just type its number (0, 1, 2, …)

**Automatically flag conflicting labels.** When two coders are labeling responses, conflicts naturally arise. These scripts will create a new column highlighting fields where a conflict occurred, and what exactly that conflict is.

**Calculate [inter-rater reliability](https://en.wikipedia.org/wiki/Inter-rater_reliability)** using [Cohen's kappa](https://en.wikipedia.org/wiki/Cohen's_kappa), [Krippendorff's alpha](https://en.wikipedia.org/wiki/Krippendorff's_alpha), or the [Kupper & Hafner metric](https://github.com/nmalkin/kupper_hafner).

**Group labels into categories** and filter out labels that you decide not to use for calculating agreement.


## Installation

Here are the steps to get started using QCAS:

1. Download and unzip the latest package [from the releases](https://github.com/nmalkin/qcas/releases)
2. In the spreadsheet where you want to use QCAS, click on _Extensions_ in the menu bar and select _Apps Script_. This will open a new window with the Apps Script Project Editor.
3. Copy over `Code.gs` from the release files
    - In the left pane, click the big + next to _Files_, and select _Script_
    - Give the file a name (e.g., `Code`) 
    - Copy and paste the contents of `Code.gs` into it.
4. Copy over `sidebar.html` from the release files
    - In the left pane, click the big + next to _Files_, and select _HTML_
    - Name the file `sidebar`
    - Copy and paste the contents of `sidebar.html` into it.
5. Though it seems intuitive, you do _not_ need to press the big blue "Deploy" button. Instead, move on to Step #6
6. Go back to the tab with your spreadsheet and refresh the page
7. After you've done this, you should see a _Coding Assistant_ menu in your menubar, to the right of the _Help_ menu.
    Select any of the menu items in that menu
8. The first time you do this, you should see a screen asking you to authorize the app you just created.
    - Despite what the prompt says, the scripts will (and can) only access the current spreadsheet.



## Setup

QCAS expects your spreadsheet to be organized a certain way.

### To resolve conflicts or calculate IRR

- Codes should be entered in two (or more) **adjacent** columns
- One row per response
- Codes should start on **row 2**. (The first row is reserved for the header.)
- If a rater assigned multiple codes to the same response, they should be in one cell, separated by commas (e.g., `code_a,code_b`)

### To take advantage of automatic renaming and categorization

This is a bit more complicated, because you need to set up two sheets: one for the codes (here, all of the above rules still apply) and one for the codebook.

- There needs to be a separate sheet for the codebook, named *question-name_codebook* with Code and Type column headers.
    - If you have only one codebook in the entire spreadsheet, you can name it `codebook` (without any prefix).
- The responses will be in a separate sheet, named *question-name_codes_coder-name* with a Coder column heading.
- If you're flagging conflicts or calculating IRR, the codes need to be pasted into a new sheet, named *question-name_codes_final*.
    - The calculations will still work if you forget to use the `_final` suffix, but the cells' background color (which is used to highlight conflicts) won't be automatically updated unless this naming scheme is followed.

[Here is a template](https://docs.google.com/spreadsheets/d/1EfjeXCM1tmDtuazxIvYET2FKieVNgM2jkuYws83mWq4) with additional explanatory comments.


## Usage

### Coding

To slightly speed up the process of coding, you can take advantage of QCAS's automatic code expansion.

1. Make sure you've set up both the coding and codebook sheets according to the instructions above
2. Select "Show codebook" from the "Coding Assistant" menu. This will open up the sidebar displaying all the codes in your codebook, with a number before each one.
3. In any cell of your coding sheet, type a number and hit Enter. It will be replaced by the associated code from your codebook.
4. To get multiple codes, type the numbers separated by spaces (e.g., `1 2 3`). After replacement, the codes will be separated by commas

For more advanced codebook operations (e.g., grouping minor codes into major ones), see further down.

### Calculating IRR

1. Select the columns with the codes for which IRR needs to be calculated
2. Go to _Coding Assistant_ > _Computer inter-rater reliability_ and select the appropriate metric

### Showing differences between two raters' codes

1. Select the columns with the codes for which differences need to be resolved
2. Go to _Coding Assistant_ > _Find conflicts_

### Type numbers instead of codes

1. Make sure you've set up the different sheets according to the instructions above
2. Make sure your codebook sheet has the codes you will be using
3. Go to _Coding Assistant_ > _Show codebook_. This will show you the numbers that will be substituted for each code.
4. For a given field, type the number of the code you want to enter. After you press _Enter_, it should be replaced with the code.
5. If you want to enter multiple codes, you can type numbers separated by spaces, e.g., `1 2` -> `code_a,code_b`

### Code renaming/grouping/categorization

1. Make sure you've set up the different sheets according to the instructions above
2. Make sure your codebook sheet has the codes you will be using and contains the `Code - final` column
3. To filter out optional codes ("flags"), enter the formula `=FILTERFLAGS(RANGE)`, where RANGE is the range where the codes you want to filter are located in A1 notation (e.g., `A1:B100`)
4. To replace codes with their new names, enter the formula `=FINALNAMES(RANGE)`, where RANGE is the range where the codes you want to rename are located in A1 notation (e.g., `A1:B100`)


## Building

This project is written in TypeScript and compiled into JavaScript before it's run. You can get the pre-compiled JS from the releases, but if you want to do it yourself:

    npm build

