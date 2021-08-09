# Awesome Excel

## Excel :skull_and_crossbones:

- [Official Excel support](https://support.office.com/en-us/excel): Every formula and function is explained
- [Official getting started guide](https://support.office.com/en-us/article/excel-for-windows-training-9bc05390-e94c-46af-a5b3-d7c22f6990bb?ui=en-US&rs=en-US&ad=US)
- [thespreadsheetguru.com](https://www.thespreadsheetguru.com/): Very good site. Also for VBA
- [Excel-Easy.com](https://www.excel-easy.com/)

### Lookups :magnet:

- [INDEX / MATCH](https://exceljet.net/index-and-match): One lookup to rule them all.
- [INDEX / MATCH Examples](https://www.excel-easy.com/examples/index-match.html): Seriously if I see one more VLookup i'm going to hit someone.
- [XLOOKUP](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929): Only works on newer Excel versions but can be used for quick lookps. Separare Index/Match is still superior in speed when pulling multiple values from the same row.

### Tables :nut_and_bolt:

Should be the Preferred way to store data, if the count of entries is variable.
You can Slice, Dice, and Reference them in a Structured way.
Also they integrate perfectly with PivotTables.

- [Official Microsoft Support](https://support.office.com/en-us/article/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c)
- [Exceljet.net Tables](https://exceljet.net/excel-tables)
- [Structured Cell References](https://support.office.com/en-us/article/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e)

## Pivot Tables :chart_with_downwards_trend: :chart_with_upwards_trend: :bar_chart:

Management friendly interactive tables and charts

- [Excel-easy.com Pivot Tables](https://www.excel-easy.com/data-analysis/pivot-tables.html)

## VBA :hammer_and_pick: :fire_extinguisher:

Automate All the things.

- [cpearson.com](http://www.cpearson.com/excel/ExcelPages.htm): is the most complete VBA Source I know. Sadly the Author passed away in 2018. Rest in Peace Chip!
- [docs.microsoft.com](https://docs.microsoft.com/en-us/office/vba/api/overview/excel): Official Object model reference for VBA. Always nice to have
- [rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck): Super Handy VBE Addin that brings some Modern IDE Features to the VBE

## Tables + VBA = :sparkling_heart:

That makes shit interesting

- [thespreadsheetguru.com](https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables): The VBA Listobject

## Power Query :gear:

Was an add-in until Excel 2010, is now in the Data Tab.
Query and Transform data from different sources. (Also used in [PowerBI](https://powerbi.microsoft.com/en-us/))
I mainly use it to get data from one or multiple other workbooks.
It basically generates Excel Tables (see above) and updates them.
Very useful, for example if you have a changing data source and want to always display all the unique entries in a column.
Also useful to transform data into a Pivot Table friendly format without changing the source data.

- [support.office.com](https://support.office.com/en-us/article/power-query-overview-and-learning-ed614c81-4b00-4291-bd3a-55d80767f81d): Official Documentation
- [docs.microsoft.com](https://docs.microsoft.com/en-us/powerquery-m/): PowerQuery M Language Reference

## Further references :telescope:

Interesting to Watch or Read

- [ebukari/awesome-excel](https://github.com/ebukari/awesome-excel)
- [Standup comedy with Spreadsheets](https://www.youtube.com/watch?v=UBX2QQHlQ_I)
