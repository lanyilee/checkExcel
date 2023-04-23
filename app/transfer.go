package app

import (
	"baliance.com/gooxml/document"
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"fyno/model"
	"fyno/utils"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
	"strings"
)

// NewTransferEntry 文件夹界面设置
func NewTransferEntry() {
	myApp := app.New()
	myWin := myApp.NewWindow("Person Transfer")

	// 文件路径
	filePath := widget.NewLabel("File Path:")
	entry1 := widget.NewEntry()
	dia1 := widget.NewButton("Open", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, myWin)
				return
			}
			if list == nil {
				log.Println("Cancelled")
				return
			}
			entry1.SetText(list.Path())
		}, myWin)
	})
	v1 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), filePath, dia1, entry1)
	execBtn := widget.NewButton("Output", func() {
		err := ReadCurrentPathFiles(entry1.Text)
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

// ReadCurrentPathFiles 读取文件夹路径下所有word文档并记录人员信息
func ReadCurrentPathFiles(path string) error {
	trans := []*model.TransferPerson{}
	if path == "" {
		path = "/Users/lanyi/documents/zzb/diaodong/2"
	}
	fInfo, err := utils.FileForEach(path)
	if err != nil {
		fmt.Println(err)
		return err
	}
	for _, o := range fInfo {
		println(o.Name())
		tp := ReadWord(path + "/" + o.Name())
		trans = append(trans, tp)
	}
	// 输出所有名字
	names := ""
	for _, t := range trans {
		if !t.IfNew {
			names += t.Name + "、"
		}

	}
	log.Println(names)

	// 输出到excel
	WriteTransferExcel(trans)
	return nil
}

func ReadWord(url string) *model.TransferPerson {
	doc, err := document.Open(url)
	//doc, err := document.Open("/Users/lanyi/Documents/zzb/diaodong/2/东组公介字【2021】175号（罗涛）.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
		return nil
	}
	//doc.Paragraphs()得到包含文档所有表格
	tp := &model.TransferPerson{}
	tp.IfNew = false

	for rowId, row := range doc.Tables()[0].Rows() {
		for cellId, cell := range row.Cells() {
			tex := ""
			for _, par := range cell.Paragraphs() {
				for _, run := range par.Runs() {
					if run.Text() == "" {
						continue
					}
					tex += run.Text()
				}
			}
			//fmt.Println("table0", "行"+strconv.Itoa(rowId), "列"+strconv.Itoa(cellId), tex)
			if rowId == 0 && cellId == 1 {
				tp.Name = tex
			} else if rowId == 1 && cellId == 1 {
				tp.Out = tex
			} else if rowId == 2 && cellId == 1 {
				tp.In = tex
			} else if rowId == 8 && cellId == 0 {
				strc := []rune(tex)
				tp.WorkDate = string(strc[5:])
			} else if rowId == 3 && cellId == 1 && tex == "/" {
				tp.IfNew = true
			}
		}
	}
	println(tp.Name + "," + tp.Out + "," + tp.In + "," + tp.WorkDate + "," + strconv.FormatBool(tp.IfNew))
	log.Println(tp.Name + "," + tp.Out + "," + tp.In + "," + tp.WorkDate + "," + strconv.FormatBool(tp.IfNew))
	return tp
}

func WriteTransferExcel(trans []*model.TransferPerson) {
	f := excelize.NewFile()
	sheetName := "Sheet1"
	//第一行
	f.SetCellStr(sheetName, "A1", "调动名单")
	f.MergeCell(sheetName, "A1", "G1")
	//第二行
	f.SetCellStr(sheetName, "A2", "序号")
	f.SetCellStr(sheetName, "B2", "姓名")
	f.SetCellStr(sheetName, "C2", "调出单位")
	f.SetCellStr(sheetName, "D2", "调入单位")
	f.SetCellStr(sheetName, "E2", "办理时间")
	f.SetCellStr(sheetName, "F2", "转任情况")
	f.SetCellStr(sheetName, "G2", "备注")

	//循环
	number := 0
	for _, t := range trans {
		nameList := strings.Split(t.Name, "、")
		nameRows := len(nameList)
		if nameRows > 1 && !t.IfNew {
			for _, k := range nameList {
				number++
				currow := strconv.Itoa(number + 2)
				f.SetCellStr(sheetName, "A"+currow, strconv.Itoa(number))
				f.SetCellStr(sheetName, "B"+currow, k)
				f.SetCellStr(sheetName, "C"+currow, t.Out)
				f.SetCellStr(sheetName, "D"+currow, t.In)
				f.SetCellStr(sheetName, "E"+currow, t.WorkDate)
				f.SetCellStr(sheetName, "F"+currow, t.Session)
				f.SetCellStr(sheetName, "G"+currow, t.Remark)
			}
		} else if !t.IfNew { //直接加一行
			number++
			currow := strconv.Itoa(number + 2)
			println("第" + currow + "行")
			f.SetCellStr(sheetName, "A"+currow, strconv.Itoa(number))
			f.SetCellStr(sheetName, "B"+currow, t.Name)
			f.SetCellStr(sheetName, "C"+currow, t.Out)
			f.SetCellStr(sheetName, "D"+currow, t.In)
			f.SetCellStr(sheetName, "E"+currow, t.WorkDate)
			f.SetCellStr(sheetName, "F"+currow, t.Session)
			f.SetCellStr(sheetName, "G"+currow, t.Remark)
		}
	}

	//设置样式
	styleId, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    2,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			Vertical:        "center",
			WrapText:        true,
		},
	})
	if err != nil {
		log.Println(err)
	}
	f.SetCellStyle(sheetName, "A1", "G"+strconv.Itoa(number+2), styleId)

	//保存
	if err := f.SaveAs("transferTest.xlsx"); err != nil {
		fmt.Println(err)
	}
}
