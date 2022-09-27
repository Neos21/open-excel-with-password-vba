# Open Excel With Password VBA

1. Open the Excel application window.
2. Press `Alt + F11` to open Visual Basic window.
3. Select from `File` menu - `Import File`.
4. Select `NeosModule.bas`. The module will be imported.
5. Select `ThisWorkbook` and write code like this:

```vb
Private Sub Workbook_Open()
  NeosModule.OpenExcelFileWithPassword "C:\PATH\TO\FILE.xlsx", "PASSWORD-HERE!"
  ThisWorkbook.Close
End Sub
```


## Links

- [Neo's World](https://neos21.net/)
