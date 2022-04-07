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


## Usage

Setting up QCAS requires a few steps.

First, the spreadsheet needs to set up in a specific way.
[Here is a template](https://docs.google.com/spreadsheets/d/1EfjeXCM1tmDtuazxIvYET2FKieVNgM2jkuYws83mWq4) with additional explanatory comments.

Note:

- There needs to be a separate sheet for the codebook, named *question-name_codebook* with Code and Type column headers.
- The responses will be in a separate sheet, named *question-name_codes_coder-name* with a Coder column heading.
- If you're flagging conflicts or calculating IRR, the codes need to be pasted into a new sheet, named *question-name_codes_final*.

Next, you need to add the scripts. This is manual and somewhat cumbersome (sorry).

In your spreadsheet, go to Extensions > Apps Script. Create new script files (by clicking the big + button under 'Files') for each of the .js files in the root of this repo (ignoring any that start with a period).
(Actually, you should be able to merge all the code into a single .js file.)

Close the script editor and refresh the spreadsheet. You should see a screen asking for permission for your scripts to run.

Once the scripts are working, you'll see a Coding Assistant menu in your menubar, to the right of the Help menu.

