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
	// シート情報を取得し、実行月のシートが無い場合は新規作成をする
	if sheet := f.GetSheetIndex(ms); sheet == -1 {
		fmt.Println("今月のシートが作成されていないため、新規にシートを作成します。")
		_ = f.NewSheet(ms)
		if err := f.SetColWidth(ms, "A", "A", 20); err != nil {
			fmt.Println(err)
			return
		}
		// 日付の欄を作成する
		for i := 1; i <= 31; i++ {
			is := strconv.Itoa(i)
			f.SetCellValue(ms, "A"+is, ms+" "+is)
		}
		// 合計値計算セルの作成
		f.SetCellValue(ms, "A32", "合計稼働時間")
		f.SetCellFormula(ms, "B32", "SUM(B1:B31)")
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

	fmt.Println("勤怠を記録しました。\n本日もお疲れ様でした。")
}
