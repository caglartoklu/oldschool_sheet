# oldschool_sheet

oldschool_sheet is a spreadsheet file with macro to convert the look of modern Microsoft Excel to the spreadsheets of DOS.

Currently, only Microsoft Excel is supported, but OpenOffice and LibreOffice support will come.

## Why created this?

First of all, I like my screen black.

Second, Excel has too much visual clutter, actually a problem of most applications of our age.
When you are working in a spreadsheet, fonts and formats and other stuff should not be more important than the content itself.
So, I wanted to strict visual features of Excel, disable font selection, row heights and color selection.

See why Game of Thrones author George R.R. Martin he writes on a DOS machine and uses
[WordStar](https://www.theverge.com/2014/5/14/5716232/george-r-r-martin-uses-dos-wordstar-to-write).

And, active cell is not recognizible in Excel.
I wanted to see if I can do something about it.

![](https://user-images.githubusercontent.com/2071639/75348295-9aef8480-58b3-11ea-8560-0930ff1ed0a7.gif)

# Requirements

- Microsoft Excel
- Tested with latest version of Office 365 at 2020-02-06, but it should work with previous versions. If it does not, please open an issue so that it can be made compatible.



# Installation

You do not need to install oldschool_sheet.
Simply download the `.xlsm` file and start using it.

See the _Usage_ section for provided functionalities.



# Usage

There are more than one way to use this, both are easy.

## 1 Use the provided `.xlsm` file and that is it.

## 2 or, Copy the macro code to your own Workbook.

To do this:

1. create a workbook on your computer with the extension of `.xlsm`.

2. open the file in Excel.

3. open _Visual Basic Editor_ by **ALT-F11**.

4. copy the contents of `this_workbook.bas` to ThisWorkbook section.

![](https://user-images.githubusercontent.com/2071639/75348311-a17dfc00-58b3-11ea-95ac-410100946a79.png)

5. create a module and put the contents of `oldschool_sheet.bas` into it.

![](https://user-images.githubusercontent.com/2071639/75348312-a17dfc00-58b3-11ea-9370-94bc943c4152.png)

![](https://user-images.githubusercontent.com/2071639/75348313-a2169280-58b3-11ea-9c7c-f722bb6c7dea.png)

6. press **CTRL-G** to open **Immediate Window**.

![](https://user-images.githubusercontent.com/2071639/75348315-a2af2900-58b3-11ea-9481-54aa329bc81b.png)

7. type `call OldSchoolMenu()` and press **enter**.

![](https://user-images.githubusercontent.com/2071639/75348316-a2af2900-58b3-11ea-9357-64aef27d6241.png)

8. done, switch to Excel main windows, click on a cell and you should see the following:

![](https://user-images.githubusercontent.com/2071639/75348318-a347bf80-58b3-11ea-8213-e12a4415d202.png)


## Using the file as a template

The `.xlsm` file can be saved as `.xltm` to be a template.

Click
[here](https://docs.microsoft.com/en-us/deployoffice/compat/office-file-format-reference#file-formats-that-are-supported-in-excel)
to read more about Excel file extension.



# Options

At the top of the module, the following options exist and they can be modified according to your needs:

```vb
Const FORMATTING_RANGE_FONT_NAME = "Consolas"
Const FORMATTING_RANGE_FONT_SIZE = 12
Const FORMATTING_RANGE_WRAP_TEXT = False
Const FORMATTING_RANGE_ROW_HEIGHT = 14.4

Const DEFAULT_RANGE = "BB200"
```

If the range formatting is too agressive, you can disable it by making it False.



# FAQ


## Is this safe?

Yes, it is safe. You can see the source code itself.

Since this is a macro, you should enable:

![](https://user-images.githubusercontent.com/2071639/75348319-a347bf80-58b3-11ea-8477-a0f525e4eb17.png)



# Development

The macro set is written in VBA (Visual Basic for Applications).

## To Do

- [x] Microsoft Excel support
- [x] Screenshots for README
- [ ] LibreOffice/OpenOffice support
- [ ] Quattro Pro color support



# License

Licensed with 2-clause license ("Simplified BSD License" or "FreeBSD License").
See the [LICENSE.txt](LICENSE.txt) file.



# Legal

All trademarks and registered trademarks are the property of their respective owners.

