package app

import (
	"errors"
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"log"
)

// MainShow 主界面函数
func MainShow() {
	//新建一个app
	a := app.New()
	//新建一个窗口
	w := a.NewWindow("近场自动化程序V1.0")
	title := widget.NewLabel("近场自动化程序")
	hello := widget.NewLabel("文件夹路径:")
	entry1 := widget.NewEntry()
	dia1 := widget.NewButton("打开", func() {
		fd := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if reader == nil {
				log.Println("Cancelled")
				return
			}
			entry1.SetText(reader.URI().Path())
		}, w)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})
	label2 := widget.NewLabel("切面方式:")
	text := widget.NewMultiLineEntry()
	text.Disable()
	//labelLast := widget.NewLabel("摩比天线技术（深圳）有限公司    ALL Right Reserved")
	labelLast := widget.NewLabel("                 ")
	combox1 := widget.NewSelect([]string{"最大值切面", "固定倾角切面"}, func(s string) { fmt.Println("selected", s) })
	label3 := widget.NewLabel("极化方式:")
	combox2 := widget.NewSelect([]string{"±45极化", "H/V极化"}, func(s string) { fmt.Println("selected", s) })
	label4 := widget.NewLabel("结果文件夹:")
	entry2 := widget.NewEntry()
	dia2 := widget.NewButton("打开", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("Cancelled")
				return
			}
			//out := fmt.Sprintf(list.String())
			entry2.SetText(list.Path())
		}, w)
	})
	combox1.SetSelectedIndex(0)
	combox2.SetSelectedIndex(0)
	bt3 := widget.NewButton("生成脚本", func() {
		if (entry1.Text != "") && (entry2.Text != "") {
			text.SetText("")
			text.Refresh()
			//txtInfo := generateTxt(entry1.Text, entry2.Text, combox2.Selected, w)
			//text.SetText("TXT脚本生成成功。请复制下面的路径信息：\n" + txtInfo)
			text.SetText(entry1.Text+entry2.Text)
			text.Refresh()
		} else {
			dialog.ShowError(errors.New("读取Excel文件错误"), w)
		}
	})
	bt4 := widget.NewButton("汇总结果", func() {
		fmt.Println(entry2.Text)
		if entry2.Text != "" {
			//bt2(entry2.Text, combox1.Selected, combox2.Selected, text)
		} else {
			dialog.ShowError(errors.New("文件夹路径错误"), w)
		}
	})
	head := container.NewCenter(title)
	v1 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), hello, dia1, entry1)
	v2 := container.NewHBox(label2, combox1)
	v3 := container.NewHBox(label3, combox2)
	v4 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), label4, dia2, entry2)
	v5 := container.NewHBox(bt3, bt4)
	v5Center := container.NewCenter(v5)
	ctnt := container.NewVBox(head, v1, v2, v3, v4, v5Center, text, labelLast)
	w.SetContent(ctnt)
	//尺寸
	w.Resize(fyne.Size{Width: 400, Height: 80})
	//w居中显示
	w.CenterOnScreen()
	//循环运行
	w.ShowAndRun()
}

