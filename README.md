# **SMG Performance Report Documentation**


**Administrators:**

**Jack Hobbs – Data Analyst –** [**jack.hobbs@papamurphys.com**](mailto:jack.hobbs@papamurphys.com)

# Introduction:

This tool was built to provide a simple way to ingest SMG store representative check-in data. It cleans the SMG report and provides a simple window into representative access history on a per-DMA basis.

VBA is a code language that runs in Microsoft&#39;s Office ecosystem. It powers the macros that are prevalent in complex projects built in Excel, PowerPoint, and Word. VBA features unique functions and methods that specifically integrate into those programs and allow a programmer to automate significant workloads, especially projects that require production of new files and datasets.

# Understanding the Analysis Tool:

The tool has two main sheets: the tracking tool interface and SMG data. The following section will cover each worksheet.

## Tracker Tool

When the user first opens the analysis tool, they are greeted with the tracker tool sheet. On this page, there are many options that the user can change to configure the output from the tool. The user interacts and sets these choices with the Slicer and Data Validation drop-down lists (DV).

### DMA Slicer:

Sets the DMAs for which the tool searches logins and returns representative info.

- Multiple options can be selected by pressing the check button at the top right of the slicer

### Representatives – Store ID:

By using the Store ID DV, the user can quickly filter to their needed store and pull up all representatives for both login stores and no-login stores

### Button – Clear:

The clear button allows the user to clear all old data from the Data tab in an instant. This allows the user to paste new data into the space without the worry of old data contaminating their most recent report. It sets the user to the correct cell to paste new info.

### Button – Run:

The Run button cleans the newly acquired SMG data by breaking out the store lists in the first column and replicating row information that is tied to that user. This will take some time (less than 2 minutes usually) and make your Microsoft applications unresponsive while it runs.

## Data:

The Data worksheet is the source from which all output tracker pivots are built. This worksheet is simple:

- Columns A:K contain the data from SMG
- Columns L:P contain calculated fields based on the SMG data
- The LL in the column titles refers to Last Login
- The LL columns convert the SMG text dates to Excel Date Values.
- &quot;True LL&quot; is the most recent login date across the desktop, mobile and app for a given store

**Remember to only paste the needed**  **data** **cells from the raw SMG file! Pasting the columns or, worse, the whole sheet into the Data worksheet will break the tool. These cells are generally found (roughly) in cells A4:K1000 of the raw SMG file.**

# VBA:

The primary intention of this section of documentation is to provide administrators with context to the code driving this tool. If an administrator needs to edit the codebase, this documentation should be an adequate guide. As with all programming, Google is the developer&#39;s best friend.

Visual Basic for Applications is a powerful tool for automating labor-intensive Word, Excel, and PowerPoint workflows. It is leveraged heavily in this analysis tool, but it is easy to edit the Excel sheet in ways that break the code. Fortunately, rectifying these edits is simple and back-ups are maintained by the administrators. VBA has been around for three decades: if a problem arises, someone on the internet has asked about it and gotten an answer.

If the developer is unfamiliar with the operation of a command, Microsoft maintains good VBA documentation: [https://docs.microsoft.com/en-us/office/vba/api/overview/](https://docs.microsoft.com/en-us/office/vba/api/overview/)

To find and open the Visual Basic Editor: [https://support.microsoft.com/en-us/office/find-help-on-using-the-visual-basic-editor-61404b99-84af-4aa3-b1ca-465bc4f45432](https://support.microsoft.com/en-us/office/find-help-on-using-the-visual-basic-editor-61404b99-84af-4aa3-b1ca-465bc4f45432)

## Definitions:

Some basic definitions, to aid the user in understanding the following documentation:

### General:

- **Module:** a place to store code, functions, variables
- **Function/Sub:** an object that can be called. When called, it executes code based on current variable values or based on values passed into it by the developer

### Variable Types:

- **String:** text data, placed in quotes.
- **Integer (Int):** numeric data, no decimals.
- **Double:** numeric data that allow decimals.
- **Workbook:** VBA variable type that represents a whole workbook.
- **Worksheet:** VBA variable type that represents a worksheet within a specified workbook.
- **Range:** VBA variable that represents an Excel range based on a given string (i.e. &quot;A1:D4&quot;)
- **Object:** VBA general variable that can represent a PowerPoint slide, file object, an Office app, etc.
- **Collection:** VBA iterable storage container like a Python list.


## Navigating the VBA Environment:

![alt text](https://github.com/Thehobbses/VBA-Analytics-Automation/blob/main/Documentation/DocumentationDiagrams/VBEditorNavigation.png)

## Module Structure:

The code for the analysis tool is in one module. The developer should be able to make most necessary changes in the SMG\_ReportTool. For the most part, functions should continue to work so long as the developer does not pass them incorrect variable types.

## Function and Method Documentation:

### SMG\_ReportTool:

#### row\_iterator:

- **Description:**

  - Connected to the Run button in the Tracker Tool worksheet, this function cleans data from the Data worksheet and breaks out rows 1:10,000 based on if the cell in the first column has commas; the delimiter used by SMG. The data loops through the rows, where it extracts comma-delimited strings of Store IDs and splits them to an array. The array is looped through and each Store ID is inserted a new row with matching data, then the ID is pasted to the first cell in that new row. Finally, the original row is deleted so only the individual Store ID rows remain.

- **Required Input Parameters:**

  - None

- **Variables:**

  - row\_index, row\_num

- **Subfunctions:**

  - None

- **Used In:**

  - None

#### Clear\_Data:

- **Description:**

  - Connected to the Clear button in the Tracker Tool worksheet, this function clears data from the Data worksheet. Specifically, it clears **all** contents in the range A2:K50000 then sets the active cell to A2 so the user can paste their new data in. Remember to only paste the needed data cells from the raw SMG file! Pasting the columns or, worse, the whole sheet into the Data worksheet will **break** the tool.

- **Required Input Parameters:**

  - None

- **Variables:**

  - row\_index, row\_num

- **Subfunctions:**

  - None

- **Used In:**

  - None

## Errors and Solutions:

Despite the best efforts of developers, VBA is a complicated programming language that imperfectly interfaces with a closed application interface. VBA errors are mercurial and can arise from many different sources. The best way to prevent these errors is to ensure the main analysis tool workbook is unaltered except where this document specifies. Most errors arise from variable type mismatches and overloaded Excel/PowerPoint application requests.

To ensure the fewest errors:

- Close all other Excel workbooks
- Close all PowerPoint presentations
- **Close and reopen Excel and try again to reset the VBA environment**
- You must close **all** Excel processes, best achieved via Task Manager

Errors will still happen, but there are best practices to resolve them. **As always, Google and Microsoft Documentation are your friends.** This section covers common errors that crop up in VBA. It is structured as follows:

Error

- Description
- Resolution

### Common Errors

**Error 429:**

- This error occurs when the user does not have the correct VBA Microsoft libraries enabled.
- Open the VBA environment and navigate to the Tools tab then References and enable the following libraries (specifically Microsoft PowerPoint 16.0 Object Library):

![alt text](https://github.com/Thehobbses/VBA-Analytics-Automation/blob/main/Documentation/DocumentationDiagrams/Error429.png)

**Error 1004 – PasteSpecial method of class Range failed:**

- The clipboard is probably empty, and the code is trying to paste nothing.
- Make sure you are passing a worksheet to Copy\_Paste\_Wait if it fails on the Excel export, or a slide if it fails on the PowerPoint export. Open and Close Excel fully and attempt again.

**Error 70 – Permission Denied:**

- This is an error that can crop up when VBA attempts to save over an open PowerPoint or Excel file.
- Reset the VBA environment by closing Excel and PowerPoint fully and opening it again. Be sure the local temporary folder or any subfiles are not open.
