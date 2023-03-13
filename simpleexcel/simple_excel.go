package simpleexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
)

func CreateExcelFileByStructSlice(dataListBySheetName map[string]any, fileName string) (err error) {
	excelFile := excelize.NewFile()
	defer func() {
		if err := excelFile.Close(); err != nil {
			panic(err)
		}
	}()

	for sheetName, data := range dataListBySheetName {
		reflectValue := reflect.ValueOf(data)
		_, err := excelFile.NewSheet(sheetName)
		if err != nil {
			panic(err)
		}

		if reflect.TypeOf(data).Kind() != reflect.Slice {
			return NewExcelError("dataList param must be struct slice", nil)
		}

		for i := 0; i < reflectValue.Len(); i++ {
			elem := reflectValue.Index(i).Interface()
			elemValue := reflect.ValueOf(elem)
			elemType := reflect.TypeOf(elem)
			if elemType.Kind() != reflect.Struct {
				return NewExcelError("dataList param must be struct slice", nil)
			}

			for j := 0; j < elemType.NumField(); j++ {
				column := string(rune('A' + j))
				if i == 0 {
					field := elemType.Field(j)
					tag := field.Tag.Get("xlsx")
					if tag == "" {
						continue
					}

					// 필드 채우기
					err = excelFile.SetCellValue(sheetName, fmt.Sprintf("%s%d", column, i+1), tag)
					if err != nil {
						return NewExcelError("err at set field name in cell", err)
					}
				}

				// 값 채우기
				err = excelFile.SetCellValue(sheetName, fmt.Sprintf("%s%d", column, i+2), elemValue.Field(j).Interface())
				if err != nil {
					return NewExcelError("err at set value in cell", err)
				}
			}
		}
	}

	buf, err := excelFile.WriteToBuffer()
	if err != nil {
		return err
	}

	err = os.WriteFile(fileName+".xlsx", buf.Bytes(), os.FileMode(0644))
	if err != nil {
		return err
	}

	return nil
}
