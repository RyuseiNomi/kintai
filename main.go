package main

import (
	"flag"
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	// 引数より、稼働時間を取得
	// ex) go run main.go 8
	flag.Parse()
	sh := flag.Arg(0)
	if sh == "" {
		fmt.Println("稼働時間を入力してください。")
		return
	}
	hour, _ := strconv.Atoi(sh)
	// 実行時の月を取得
	t := time.Now()
	m := t.Month()
	ms := m.String()
	d := t.Day()
	ds := strconv.Itoa(d)
	log.Println(ms, ds)
	// Excelファイルを開く
	f, err := excelize.OpenFile("./sheet/kintai.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// シートのindex番号を取得
	index := f.GetSheetIndex(ms)
	// 値をSet
	f.SetCellValue(ms, "B"+ds, hour)
	f.SetActiveSheet(index)
	// 保存
	if err := f.Save(); err != nil {
		fmt.Println(err)
		return
	}
	// 合計稼働時間の取得
	th, err := f.GetCellValue(ms, "B32")
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("勤怠を記録しました。\n今月の合計稼働時間は、現在 " + th + " 時間です。\n本日もお疲れ様でした。")
}
