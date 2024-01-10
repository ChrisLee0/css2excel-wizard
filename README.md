# Css2Excel-Wizard: Generate Excel with CSS-like Syntax

Css2Excel-Wizard is a Java utility that simplifies the process of generating Excel files by allowing you to utilize CSS-like syntax for styling, as well as conveniently set up merged cells and spanning rows/columns. With this tool, you can effortlessly create visually appealing and structurally organized Excel spreadsheets.

## Key Features:

* CSS-style Styling: Css2Excel-Wizard enables you to apply styles to Excel cells using a familiar CSS-like syntax. You can specify properties such as font, color, alignment, borders, background, and more, making it intuitive and straightforward to achieve the desired visual appearance for your data.

* Merged Cells: This tool provides easy-to-use functions to merge cells in Excel, allowing you to combine multiple cells horizontally or vertically. You can span cells across rows or columns effortlessly, enabling the creation of complex layouts and organized data structures.

* Flexible Configuration: Css2Excel-Wizard offers a range of configuration options to fine-tune your Excel output. You can customize the sheet name, define the starting cell for data insertion, set the width of columns, specify the sheet orientation, and more. These configurations provide control over the overall structure and layout of the generated Excel file.

* Efficient and Lightweight: Built with efficiency in mind, Css2Excel-Wizard is designed to generate Excel files quickly and consume minimal system resources. It offers a streamlined approach to Excel file generation, making it an ideal choice for projects requiring high-performance spreadsheet creation.

## Showcase:

File: **Streamlined Variable Pay.xlsx**

![image](https://github.com/ChrisLee0/css2excel-wizard/blob/master/img/SVP.png)

```java
Css2ExcelWizard css2ExcelWizard = new Css2ExcelWizard();
css2ExcelWizard.createSheet("SVP");
css2ExcelWizard.getLayout().setWidth(2, 12, 8, 8, 8, 10); //define root class, all class will default inherit properties from here
css2ExcelWizard.getCss().defineRootStyle("font-family: Calibri; font-size: 10;border-color: black;;vertical-align: middle");

//define a class called 'cell'
css2ExcelWizard.getCss().defineStyle("cell", "text-align: center");
//define a class called 'heading' which inherits properties from class 'cell'
css2ExcelWizard.getCss().extendStyle("cell", "heading", "font-weight: bold; font-size: 14;")
		.extendStyle("cell", "header", "background-color: red; color: white; font-weight: bold; ")
		//defined a style called 'rate' with a number with 2 decimals
		.extendStyle("cell", "rate", "number-format: 0.00").extendStyle("cell", "grey-cell", "background-color: rgb(191,191,191)");
css2ExcelWizard.getCss().defineStyle("note", "text-align: left;border: none; wrap-text: true");
//defined a style with a formatted date
css2ExcelWizard.getCss().defineStyle("date", "text-align: left;border: none; font-weight: bold;date-format: dd-mmm-yyyy");

css2ExcelWizard.getLayout().newRow()
	.newRow().cell().cellSpan(1, 5, "heading", "Streamlined Variable Pay")
	.newRow().cell().cell("header", "Ratings", "Top", "Strong", "Good", "Inconsistent")
		.newRow().cell().cell("header", "Role Model").cell("rate", 3, 2, 1.25).cell("grey-cell")
		.newRow().cell().cell("header", "Strong").cell("rate", 2.5, 1.5, 1).cell("grey-cell")
		.newRow().cell().cell("header", "Good").cell("rate", 2.25, 1.25, 0.75).cell("grey-cell")
		.newRow().cell().cell("header", "Unacceptable").cell("grey-cell").cell("grey-cell").cell("grey-cell").cell("grey-cell").newRow()
		//create a row with height 40, which takes 1 row and 7 cols, a rich text inside. //grammar: normal text {fontId|text....} normal text
		.newRow(40).cell().cellSpan(1, 7, "note", css2ExcelWizard.getCss().richText("As an example using the grid shown, an employee rated {good|“good performer”} for performance and {str|“strong”} for behaviours will receive {emp|one month's fixed pay as their variable pay}.",
				css2ExcelWizard.getLayout().cssManager.newRichTexTLocalStyle().defineDefaultFont("10 black normal").defineFont("good", "12 orange bold").defineFont("str", "11 red bold").defineFont("emp", "11 green double"))).newRow().newRow().cell().cell("date", new Date());


//Important: do not forget to call appendMergedCells at the end if you have merged cell
css2ExcelWizard.getLayout().appendMergedCells();
//hide grid lines
css2ExcelWizard.setDisplayGridLines(false);
css2ExcelWizard.createSheet("Test");
css2ExcelWizard.getLayout().newRow().cellSpan(4, 4, "heading", "Merged Cell").cell(null, 1, 2, 3, 4).newRow().cell(null, "New line start from here").newRow().cell().cellSpan(4, 2, "heading", "Merged Cell").newRow().cell(null, "New line start from here").newRow().newRow().newRow();

css2ExcelWizard.getLayout().autoWidthFitsContent(4);
css2ExcelWizard.getLayout().appendMergedCells();

css2ExcelWizard.writeFile("./Streamlined Variable Pay.xlsx");

```
