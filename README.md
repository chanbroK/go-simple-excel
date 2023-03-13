# simple excel read/write code
> Use [Excelize](https://github.com/qax-os/excelize) package

## Write
- write excel sheet and file by using go slice struct
- Example
```go
SheetDataList := make([]SheetData, 0)
for i := 1; i <= 10; i++ {
	SheetDataList = append(SheetDataList, SheetData{
		Col1: fmt.Sprintf("A%d", i),
		Col2: fmt.Sprintf("B%d", i),
		Col3: i * 10,
		Col4: 0.001 * float64(i),
		Col5: fmt.Sprintf("C%d", i),
		Col6: fmt.Sprintf("D%d", i),
	})
}
SheetNameBySheetDataList := map[string]any{
	"Sheet1": SheetDataList,
	"Sheet2": SheetDataList,
	"Sheet3": SheetDataList,
}
err := simpleexcel.CreateExcelFileByStructSlice(SheetNameBySheetDataList, "test")
if err != nil {		
	panic(err)
}
```
## Read
