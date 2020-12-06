package main

import (
	"log"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	// Excelファイルを開く
	f, err := excelize.OpenFile("./sheet/kintai.xlsx")
	if err != nil {
		log.Println(err)
		return
	}
	// シートのindex番号を取得
	index := f.GetSheetIndex("12月")
	// 値をSet
	f.SetCellValue("12月", "B2", 8)
	f.SetActiveSheet(index)
	// 保存
	if err := f.Save(); err != nil {
		log.Println(err)
		return
	}
	log.Println("completed to record your work record!")
}
