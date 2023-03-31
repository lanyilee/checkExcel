package app

import (
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"fyno/model"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
)

func NewTestEntry() {
	myApp := app.New()
	myWin := myApp.NewWindow("Excel Test")

	// excel表1
	excelPath := widget.NewLabel("Excel Path:")
	entry1 := widget.NewEntry()
	dia1 := widget.NewButton("Open", func() {
		fd := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, myWin)
				return
			}
			if reader == nil {
				log.Println("Cancelled")
				return
			}
			entry1.SetText(reader.URI().Path())
		}, myWin)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})
	v1 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), excelPath, dia1, entry1)
	execBtn := widget.NewButton("Output", func() {
		err := ReadTestExcel(entry1.Text)
		if err != nil {
			log.Println(err)
		}
	})

	labelLast := widget.NewLabel("                 ")
	content := container.NewVBox(v1, execBtn, labelLast)

	myWin.SetContent(content)
	myWin.Resize(fyne.NewSize(800, 600))
	myWin.ShowAndRun()
}

// ReadTestExcel 读取test考核excel
func ReadTestExcel(path string) error {
	if path == "" {
		path = "/Users/lanyi/documents/zzb/test.xlsx"
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return err
	}
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return err
	}
	// 读取单位信息
	curUnit := model.Unit{}
	for rowIndex, row := range rows {
		if rowIndex < 3 {
			continue
		}
		//单位信息
		up := model.UnitPerson{}
		for colIndex, colCell := range row {
			switch colIndex {
			case 0:
				up.UnitName = colCell
			case 5:
				up.FirstQuarter = colCell
			case 6:
				up.FirstQuarterRecord = colCell
			case 7:
				up.SecondQuarter = colCell
			case 8:
				up.SecondQuarterRecord = colCell
			case 9:
				up.ThirdQuarter = colCell
			case 10:
				up.ThirdQuarterRecord = colCell
			case 11:
				up.FourthQuarter = colCell
			case 12:
				up.FourthQuarterRecord = colCell
			case 15:
				up.Type = colCell
			}
		}
		// 判断当前行和当前计算单位是否同一个
		if up.UnitName != curUnit.UnitName {
			// 旧的先结算
			if curUnit.UnitName != "" {
				CountUnit(curUnit)
			}
			// 新的初始化
			curUnit = model.Unit{}
			curUnit.UnitName = up.UnitName
		}
		// 人员类别和4个季度备注都无则纳入计算
		if up.FirstQuarterRecord == "" && up.SecondQuarterRecord == "" && up.
			ThirdQuarterRecord == "" && up.FourthQuarterRecord == "" && up.Type == "" {
			if up.FirstQuarter == "好" {
				curUnit.FirstGood++
			} else if up.FirstQuarter == "较好" || up.FirstQuarter == "不确定等次" {
				curUnit.FirstAboutGood++
			}
			if up.SecondQuarter == "好" {
				curUnit.SecondGood++
			} else if up.SecondQuarter == "较好" || up.SecondQuarter == "不确定等次" {
				curUnit.SecondAboutGood++
			}
			if up.ThirdQuarter == "好" {
				curUnit.ThirdGood++
			} else if up.ThirdQuarter == "较好" || up.ThirdQuarter == "不确定等次" {
				curUnit.ThirdAboutGood++
			}
			if up.FourthQuarter == "好" {
				curUnit.FourthGood++
			} else if up.FourthQuarter == "较好" || up.FourthQuarter == "不确定等次" {
				curUnit.FourthAboutGood++
			}
		}
	}
	//最后一个计算
	CountUnit(curUnit)
	// 关闭工作簿
	if err = f.Close(); err != nil {
		fmt.Println(err)
	}
	return nil
}

func CountUnit(curUnit model.Unit) {
	//第一季度
	firstSum := curUnit.FirstAboutGood + curUnit.FirstGood
	firstPercent := 0.0
	if firstSum > 0 {
		firstPercent, _ = strconv.ParseFloat(fmt.Sprintf("%.2f",
			float32(curUnit.FirstGood)/float32(firstSum)), 64)
	}
	//二
	secondSum := curUnit.SecondAboutGood + curUnit.SecondGood
	secondPercent := 0.0
	if secondSum > 0 {
		secondPercent, _ = strconv.ParseFloat(fmt.Sprintf("%.2f",
			float32(curUnit.SecondGood)/float32(secondSum)), 64)
	}
	//三
	thirdSum := curUnit.ThirdAboutGood + curUnit.ThirdGood
	thirdPercent := 0.0
	if thirdSum > 0 {
		thirdPercent, _ = strconv.ParseFloat(fmt.Sprintf("%.2f",
			float32(curUnit.ThirdGood)/float32(thirdSum)), 64)
	}
	//四
	fourthSum := curUnit.FourthAboutGood + curUnit.FourthGood
	fourthPercent := 0.0
	if fourthSum > 0 {
		fourthPercent, _ = strconv.ParseFloat(fmt.Sprintf("%.2f",
			float32(curUnit.FourthGood)/float32(fourthSum)), 64)
	}
	if firstPercent > 0.4 || secondPercent > 0.4 || thirdPercent > 0.4 || fourthPercent > 0.4 {
		curUnit.IfPrint = true
	}
	if curUnit.IfPrint {
		println(curUnit.UnitName + " 超过40%；")
		log.Println(curUnit.UnitName + " 超过40%；")
	}
}
