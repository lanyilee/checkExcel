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
	"log"
	"strings"
)

// NewCompareEntry 文件夹界面设置
func NewCompareEntry() {
	myApp := app.New()
	myWin := myApp.NewWindow("Compare Person")

	// 文件夹1路径(公统)
	filePath1 := widget.NewLabel("Public File Path:")
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
	v1 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), filePath1, dia1, entry1)
	// 文件夹2路径（专项）
	filePath2 := widget.NewLabel("Special File Path:")
	entry2 := widget.NewEntry()
	dia2 := widget.NewButton("Open", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, myWin)
				return
			}
			if list == nil {
				log.Println("Cancelled")
				return
			}
			entry2.SetText(list.Path())
		}, myWin)
	})
	v2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), filePath2, dia2, entry2)
	execBtn := widget.NewButton("Output", func() {
		//println(entry1.Text)
		//println(entry2.Text)
		err := ReadComparePathFiles(entry1.Text, entry2.Text)
		if err != nil {
			log.Println(err)
		}
	})

	labelLast := widget.NewLabel("                 ")
	content := container.NewVBox(v1, v2, execBtn, labelLast)

	myWin.SetContent(content)
	myWin.Resize(fyne.NewSize(800, 600))
	myWin.ShowAndRun()
}

// ReadComparePathFiles 读取文件夹(公统)路径下所有word文档
func ReadComparePathFiles(publicPath string, specialPath string) error {
	// 记录specialPath所有文件名
	fSInfo, err := utils.FileForEach(specialPath)
	if err != nil {
		fmt.Println(err)
		return err
	}
	specialFileNames := []string{}
	for _, o := range fSInfo {
		//println(o.Name())
		specialFileNames = append(specialFileNames, o.Name())
	}
	// 记录publicPath所有文件名
	fInfo, err := utils.FileForEach(publicPath)
	if err != nil {
		fmt.Println(err)
		return err
	}
	publicNames := []string{}
	for _, o := range fInfo {
		//公统名字位0000_xxx.doc类型
		nameByte := []rune(o.Name())
		nameLen := len(nameByte)
		name := ""
		if nameLen > 10 {
			name = string(nameByte[5 : nameLen-5])
		} else {
			continue
		}
		if name == "" {
			continue
		}
		//println(name)
		// 遍历专项
		flag := false
		for _, spfile := range specialFileNames {
			// 有相同名字，作对比
			if strings.Contains(spfile, name) {
				publicUrl := publicPath + "/" + o.Name()
				specialUrl := specialPath + "/" + spfile
				CompareTwoWord(publicUrl, specialUrl)
				flag = true
			}
		}
		// 如果专项查不到，看是转任还是考录
		if !flag {
			publicUrl := publicPath + "/" + o.Name()
			err = GetPublicWord(publicUrl)
			if err != nil {
				return err
			}
		}
		publicNames = append(publicNames, name)
	}
	return nil
}

// CompareTwoWord 对比两个word文档内容
func CompareTwoWord(publicUrl string, specialUrl string) error {
	// 先打开公统表
	publicDoc, err := document.Open(publicUrl)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
		return err
	}
	//doc.Paragraphs()得到包含文档所有表格
	cp1 := &model.ComparePerson{}
	for rowId, row := range publicDoc.Tables()[0].Rows() {
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
				cp1.Name = tex
			} else if rowId == 0 && cellId == 3 {
				cp1.Sex = tex
			} else if rowId == 0 && cellId == 5 {
				if len(tex) >= 7 {
					cp1.Birth = tex[:7]
				} else {
					log.Println(cp1.Name + "：出生时间有误")
				}
			} else if rowId == 1 && cellId == 1 {
				cp1.Nation = utils.RemoveStringChineseSpace(tex)
				//strc := []rune(tex)
				//tp.WorkDate = string(strc[5:])
			} else if rowId == 1 && cellId == 3 {
				cp1.NationPlace = tex
			} else if rowId == 1 && cellId == 5 {
				cp1.BirthPlace = tex
			} else if rowId == 2 && cellId == 3 {
				cp1.JoinInWork = tex
			} else if rowId == 2 && cellId == 5 {
				cp1.IfHealth = utils.RemoveStringChineseSpace(tex)
			} else if rowId == 4 && cellId == 2 {
				cp1.FullTimeEducation = tex
			} else if rowId == 4 && cellId == 4 {
				cp1.FullTimeSchool = tex
			} else if rowId == 9 && cellId == 1 {
				cp1.Resume = tex
			}
		}
	}
	println(cp1.Name + "," + cp1.Sex + "," + cp1.Nation + "," + cp1.NationPlace + "," + cp1.Birth + "," +
		cp1.BirthPlace + "," + cp1.IfHealth + "," + cp1.FullTimeEducation + "," + cp1.FullTimeSchool + "," +
		cp1.JoinInWork + "," + cp1.Resume)
	//log.Println(cp1.Name + "," + cp1.Sex + "," + cp1.Nation + "," + cp1.NationPlace + "," + cp1.Birth+ "," +
	//	cp1.BirthPlace+ "," +cp1.IfHealth+ "," +cp1.FullTimeEducation+ "," +cp1.FullTimeSchool+ "," +
	//	cp1.JoinInWork+ "," +cp1.Resume)

	// 打开专项表
	specialDoc, err := document.Open(specialUrl)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
		return err
	}
	//doc.Paragraphs()得到包含文档所有表格
	cp2 := &model.ComparePerson{}
	for rowId, row := range specialDoc.Tables()[0].Rows() {
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
			//fmt.Println("table1", "行"+strconv.Itoa(rowId), "列"+strconv.Itoa(cellId), tex)
			if rowId == 0 && cellId == 1 {
				cp2.Name = tex
			} else if rowId == 0 && cellId == 3 {
				cp2.Sex = tex
			} else if rowId == 0 && cellId == 5 {
				if len(tex) >= 7 {
					cp2.Birth = tex[:7]
				} else {
					log.Println(cp1.Name + "：出生时间有误")
				}
			} else if rowId == 1 && cellId == 1 {
				cp2.Nation = utils.RemoveStringChineseSpace(tex)
			} else if rowId == 1 && cellId == 3 {
				cp2.NationPlace = tex
			} else if rowId == 1 && cellId == 5 {
				cp2.BirthPlace = tex
			} else if rowId == 2 && cellId == 3 {
				cp2.JoinInWork = tex
			} else if rowId == 2 && cellId == 5 {
				cp2.IfHealth = utils.RemoveStringChineseSpace(tex)
			} else if rowId == 4 && cellId == 2 {
				cp2.FullTimeEducation = tex
			} else if rowId == 4 && cellId == 4 {
				cp2.FullTimeSchool = tex
			} else if rowId == 9 && cellId == 1 {
				cp2.Resume = tex
			}
		}
	}
	println(cp2.Name + "," + cp2.Sex + "," + cp2.Nation + "," + cp2.NationPlace + "," + cp2.Birth + "," +
		cp2.BirthPlace + "," + cp2.IfHealth + "," + cp2.FullTimeEducation + "," + cp2.FullTimeSchool + "," +
		cp2.JoinInWork + "," + cp2.Resume)

	// 对比
	logMsg := cp1.Name
	if strings.Compare(cp1.Sex, cp2.Sex) != 0 {
		logMsg += " 性别有错,专项为" + cp2.Sex + "," + "公统为" + cp1.Sex + ";"
	}
	if strings.Compare(cp1.Birth, cp2.Birth) != 0 {
		logMsg += " 出生年月有错,专项为" + cp2.Birth + "," + "公统为" + cp1.Birth + ";"
	}
	if strings.Compare(cp1.Nation, cp2.Nation) != 0 {
		logMsg += " 民族有错,专项为" + cp2.Nation + "," + "公统为" + cp1.Nation + ";"
	}
	if strings.Compare(cp1.NationPlace, cp2.NationPlace) != 0 {
		logMsg += " 籍贯有错,专项为" + cp2.NationPlace + "," + "公统为" + cp1.NationPlace + ";"
	}
	if strings.Compare(cp1.BirthPlace, cp2.BirthPlace) != 0 {
		logMsg += " 出生地有错,专项为" + cp2.BirthPlace + "," + "公统为" + cp1.BirthPlace + ";"
	}
	if strings.Compare(cp1.IfHealth, cp2.IfHealth) != 0 {
		logMsg += " 健康状况有错,专项为" + cp2.BirthPlace + "," + "公统为" + cp1.BirthPlace + ";"
	}
	if strings.Compare(cp1.JoinInWork, cp2.JoinInWork) != 0 {
		logMsg += " 参加工作时间有错,专项为" + cp2.JoinInWork + "," + "公统为" + cp1.JoinInWork + ";"
	}
	if strings.Compare(cp1.FullTimeEducation, cp2.FullTimeEducation) != 0 {
		logMsg += " 全日制教育有错,专项为" + cp2.FullTimeEducation + "," + "公统为" + cp1.FullTimeEducation + ";"
	}
	if strings.Compare(cp1.FullTimeSchool, cp2.FullTimeSchool) != 0 {
		logMsg += " 全日制学校有错,专项为" + cp2.FullTimeSchool + "," + "公统为" + cp1.FullTimeSchool + ";"
	}
	logMsg += "\n"
	println(logMsg)
	curPath := utils.GetCurrentAbPath() + "/compare_add.txt"
	err = utils.WriteText(curPath, logMsg)
	if err != nil {
		fmt.Println(err)
		log.Println(err)
		return err
	}
	return nil
}

func GetPublicWord(publicUrl string) error {
	// 打开公统表
	publicDoc, err := document.Open(publicUrl)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
		return err
	}
	cp1 := &model.ComparePerson{}
	for rowId, row := range publicDoc.Tables()[0].Rows() {
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
				cp1.Name = tex
			} else if rowId == 0 && cellId == 3 {
				cp1.Sex = tex
			} else if rowId == 0 && cellId == 5 {
				if len(tex) >= 7 {
					cp1.Birth = tex[:7]
				} else {
					log.Println(cp1.Name + "：出生时间有误")
				}
			} else if rowId == 1 && cellId == 1 {
				cp1.Nation = utils.RemoveStringChineseSpace(tex)
			} else if rowId == 1 && cellId == 3 {
				cp1.NationPlace = tex
			} else if rowId == 1 && cellId == 5 {
				cp1.BirthPlace = tex
			} else if rowId == 2 && cellId == 3 {
				cp1.JoinInWork = tex
			} else if rowId == 2 && cellId == 5 {
				cp1.IfHealth = utils.RemoveStringChineseSpace(tex)
			} else if rowId == 4 && cellId == 2 {
				cp1.FullTimeEducation = tex
			} else if rowId == 4 && cellId == 4 {
				cp1.FullTimeSchool = tex
			} else if rowId == 9 && cellId == 1 {
				cp1.Resume = tex
			}
		}
	}
	logMsg := "新增人员：" + cp1.Name + "," + cp1.Sex + "," + cp1.Nation + "," + cp1.NationPlace + "," + cp1.Birth + "," +
		cp1.BirthPlace + "," + cp1.IfHealth + "," + cp1.FullTimeEducation + "," + cp1.FullTimeSchool + "," +
		cp1.JoinInWork + "," + cp1.Resume + "\n"
	curPath := utils.GetCurrentAbPath() + "/compare.txt"
	utils.WriteText(curPath, logMsg)
	return nil
}
