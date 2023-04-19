package main

import (
	"encoding/json"
	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
	"os"
)

// Excel配置
type ExcelConf struct {
	DataFile   string  `json:"DataFile"`
	SheetName  string  `json:"SheetName"`
	Fields     []Field `json:"Fields"` // 字段-列号映射
	OutputPath string  `json:"OutputPath"`
}

type Field struct {
	FieldName string
	FieldPath string
}

// 读取Excel配置
func readExcelConfig() ExcelConf {
	data, _ := os.ReadFile("conf.json")
	var conf ExcelConf
	json.Unmarshal(data, &conf)
	return conf
}

func main() {
	// 读取配置
	conf := readExcelConfig()

	// 解析JSON
	// 读取data.json
	dataJSON, _ := os.ReadFile(conf.DataFile)

	// 新建Excel
	f := excelize.NewFile()
	_, err := f.NewSheet(conf.SheetName)
	if err != nil {
		panic("创建sheet失败")
	}

	for col, field := range conf.Fields {
		cell, err := excelize.CoordinatesToCellName(col+1, 1)
		if err != nil {
			panic(err)
		}
		f.SetCellValue(conf.SheetName, cell, field.FieldName)
	}

	// 插入数据
	colIndex := 0
	gjson.ParseBytes(dataJSON).ForEach(func(key, value gjson.Result) bool {
		colIndex++
		for col, field := range conf.Fields {

			cell, _ := excelize.CoordinatesToCellName(col+1, colIndex+1)
			excelValue := gjson.Get(value.String(), field.FieldPath).String()

			f.SetCellValue(conf.SheetName, cell, excelValue)
		}
		return true
	})
	// 保存Excel
	f.SaveAs(conf.OutputPath)
}
