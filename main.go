package main

import (
	"fmt"
	"go-simple-excel/simpleexcel"
)

func main() {
	type SheetData struct {
		Col1 string  `xlsx:"col1"`
		Col2 string  `xlsx:"col2"`
		Col3 int     `xlsx:"col3"`
		Col4 float64 `xlsx:"col4"`
		Col5 string  `xlsx:"col5"`
		Col6 string  `xlsx:"col6"`
	}
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

}
