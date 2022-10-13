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
	"fyno/constant"
	"fyno/model"
	"fyno/utils"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
	"strings"
	"time"
)

func NewExcelEntry() {
	myApp := app.New()
	myWin := myApp.NewWindow("Excel 统计工具")

	//nameEntry := widget.NewEntry()
	//nameEntry.SetPlaceHolder("Please input number")
	////设置只能输入数字
	//nameEntry.Validator=validation.NewRegexp("^[0-9]+$","Please input number")
	//nameEntry.OnChanged = func(content string) {
	//	fmt.Println("name:", nameEntry.Text, "entered")
	//}
	////nameEntry.Wrapping=fyne.TextWrapBreak
	//nameBox := container.NewVBox(widget.NewLabel("Rows:"), nameEntry)

	//path
	//pathEntry := widget.NewEntry()
	//pathEntry.SetPlaceHolder("Please input excel project path")
	//pathBox := container.NewVBox(widget.NewLabel("Path:"), pathEntry)

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

	//excel表2
	excel2Path := widget.NewLabel("BigDataExcel Path大数据:")
	entry2 := widget.NewEntry()
	dia2 := widget.NewButton("Open", func() {
		fd := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, myWin)
				return
			}
			if reader == nil {
				log.Println("Cancelled")
				return
			}
			entry2.SetText(reader.URI().Path())
		}, myWin)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})
	v2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), excel2Path, dia2, entry2)

	bt2 := widget.NewButton("Output", func() {
		fmt.Println(entry1.Text)
		if entry1.Text != "" && entry2.Text != "" {
			cities, err := ReadExcel(entry1.Text)
			if err != nil {
				println(err)
				log.Println(err)
				dialog.ShowError(err, myWin)
			}
			WriteExcel(cities, entry2.Text)
		} else {
			dialog.ShowError(errors.New("error path"), myWin)
		}
	})
	button2 := container.NewHBox(bt2)
	button2Center := container.NewCenter(button2)
	text := widget.NewMultiLineEntry()
	text.Disable()
	labelLast := widget.NewLabel("                 ")
	content := container.NewVBox(v1, v2, button2Center, text, labelLast)
	//execBtn := widget.NewButton("Start", func() {
	//
	//	fmt.Println("rows:", nameEntry.Text, "Start")
	//	ReadExcel(pathEntry.Text)
	//})

	//content := container.NewVBox(nameBox,pathBox,layout.NewSpacer(), execBtn)

	myWin.SetContent(content)
	myWin.Resize(fyne.NewSize(800, 600))
	myWin.ShowAndRun()
}

// ReadExcel 读取深圳等重点地区返骆人员登记表.xlsx，固定格式
func ReadExcel(path string) ([]model.ProvinceCity, error) {
	if path == "" {
		path = "/Users/lanyi/documents/防疫/test.xlsx"
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return nil, err
	}
	//城市统计数据
	cityMap := map[string]model.ProvinceCity{}
	//其他不能处理的数据
	others := []model.Person{}

	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("今日排查")
	if err != nil {
		fmt.Println(err)
		return nil, err
	}
	for rowIndex, row := range rows {
		if rowIndex < 2 {
			continue
		}
		//人员信息
		person := model.Person{}
		for colIndex, colCell := range row {
			switch colIndex {
			case 0:
				person.Number = colCell
			case 1:
				person.Name = colCell
			case 2:
				person.Sex = colCell
			case 3:
				person.Age = colCell
			case 4:
				person.Id = colCell
			case 5:
				person.Phone = colCell
			//计数村
			case 6:
				person.Village = colCell
			case 7:
				person.LeaveTime = colCell
			case 8:
				person.BackTime = colCell
			case 9:
				person.Control = colCell
			case 10:
				person.ControlTime = colCell
			case 11:
				person.PrevAddr = colCell
			}
			//fmt.Print(colCell, "\t")
		}
		selectCity := ""
		//遍历城市表
		for cityIndex, city := range constant.Cities {
			if strings.Contains(person.PrevAddr, city) {
				//判断hash表有无当前城市，无则添加hash,有则跳过
				if _, ok := cityMap[city]; !ok {
					cityMap[city] = model.ProvinceCity{
						CityName: city,
						Index:    cityIndex,
					}
				}
				selectCity = city
				break
			}
		}
		//判断表格城市是否在已定城市列表中，不在的话，直接返回该条数据
		if selectCity == "" {
			others = append(others, person)
			continue
		}
		//遍历村居表
		curCity := cityMap[selectCity] //从hash表获取当前city
		flag := false                  //找到村居企业退出循环标志
		for _, village := range constant.Villages {
			if strings.Contains(person.Village, village) {
				for i, v := range curCity.Villages {
					//有表格中的村居数据，++
					if v.Name == village {
						curCity.Villages[i].Number++
						curCity.Number++
						//hash表重新赋值
						cityMap[selectCity] = curCity
						flag = true //标志找到村居，可以退出大循环
						break
					}
				}
				if flag {
					break
				}
				//判断当前城市实体中有无当前村居的计数，无则添加,有则+1
				vill := model.Village{}
				vill.Name = village
				vill.Number = 1
				curCity.Villages = append(curCity.Villages, vill)
				curCity.Number++
				flag = true
				//hash表重新赋值
				cityMap[selectCity] = curCity
				break
				//if len(curCity.Villages)==0{
				//}
			}
		}

		//村居或企业在已知村居数组中找不到，就把原数据返回
		if !flag {
			others = append(others, person)
		}
	}
	// 关闭工作簿
	if err = f.Close(); err != nil {
		fmt.Println(err)
	}
	//统计表遍历,打印日志部分
	//分类统计，省内（除深圳），外省
	guangdong := 0
	waisheng := 0
	guangdongStr := ""
	waishengStr := ""
	//省外村居
	foreignVilleage := map[string]int{}
	foreignVilleageArray := []string{}
	//城市数组,最后加省外，无论省外有无数据
	citiesArray := []model.ProvinceCity{}
	for cityIndex, cityName := range constant.Cities {
		if _, ok := cityMap[cityName]; ok {
			city := cityMap[cityName]
			if cityIndex == 1 { //深圳
				guangdongStr += city.CityName + strconv.Itoa(city.Number) + "人，"
				viStr := "\n" + city.CityName + "总数为 " + strconv.Itoa(city.Number) + ",其中："
				for _, v := range city.Villages {
					viStr += v.Name + "为 " + strconv.Itoa(v.Number) + " ;"
				}
				//add to cities array
				citiesArray = append(citiesArray, city)
				log.Println(viStr)
				println(viStr)
			} else if cityIndex != 1 && cityIndex < 20 { //省内
				guangdong += city.Number
				guangdongStr += city.CityName + strconv.Itoa(city.Number) + "人，"
				viStr := "\n" + city.CityName + "总数为 " + strconv.Itoa(city.Number) + ",其中："
				for _, v := range city.Villages {
					viStr += v.Name + "为 " + strconv.Itoa(v.Number) + " ;"
				}
				//add to cities array
				citiesArray = append(citiesArray, city)
				println(viStr)
				log.Println(viStr)
			} else { //省外
				waisheng += city.Number
				waishengStr += city.CityName + strconv.Itoa(city.Number) + "人，"
				//省外对应村居
				for _, v := range city.Villages {
					if _, ok := foreignVilleage[v.Name]; ok {
						foreignVilleage[v.Name] += v.Number
						continue
					}
					foreignVilleage[v.Name] = v.Number
					foreignVilleageArray = append(foreignVilleageArray, v.Name)
				}
				print("\n" + city.CityName + "总数为 " + strconv.Itoa(city.Number))
			}
			println("")
		}
	}
	log.Println("省内(除深圳)总人数：" + strconv.Itoa(guangdong))
	log.Println(guangdongStr)

	//遍历外省对应村居
	foreignProvince := model.ProvinceCity{}
	foreignProvince.CityName = "省外"
	foreignProvince.Number = waisheng
	log.Println("外省总人数：" + strconv.Itoa(waisheng) + "，其中:")
	for _, v := range foreignVilleageArray {
		village := model.Village{}
		village.Name = v
		village.Number = foreignVilleage[v]
		foreignProvince.Villages = append(foreignProvince.Villages, village)
		waishengStr += v + "为 " + strconv.Itoa(foreignVilleage[v]) + ";"
		print(v + "为 " + strconv.Itoa(foreignVilleage[v]) + ";")
	}
	//省外对象加入到城市数组中
	citiesArray = append(citiesArray, foreignProvince)

	log.Println(waishengStr)
	println("\n" + "省内总人数：" + strconv.Itoa(guangdong))
	println("外省总人数：" + strconv.Itoa(waisheng))
	//其他城市或村居遍历
	println("\n" + "其他数据：")
	log.Println("\n" + "其他数据：")
	for _, v := range others {
		log.Println(v.Number + " " + v.Name + " " + v.Phone + " " + v.PrevAddr + " " + v.Village + ";")
		println(v.Number + " " + v.Name + " " + v.Phone + " " + v.PrevAddr + " " + v.Village + ";")
	}
	//返回数组
	return citiesArray, nil
}

// TestWriteExcel 测试输出excel表格
func TestWriteExcel() {
	f := excelize.NewFile()
	// 创建一个工作表
	index := f.NewSheet("Sheet2")
	f.SetCellValue("sheet1", "B2", "LiuBei")
	f.SetCellValue("sheet1", "B9", "GuanYu")
	f.MergeCell("Sheet1", "B2", "D4") //合并单元格的对角线单元格坐标
	f.MergeCell("Sheet1", "B9", "E10")
	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

// WriteExcel 输出excel表格
func WriteExcel(cities []model.ProvinceCity, path string) {

	f := excelize.NewFile()

	sheetName := "Sheet1"
	WriteHead(f, cities, sheetName)
	villageNames := WriteBodyLeft(f, cities, sheetName)
	println(villageNames)
	WriteBodyRight(f, sheetName, path, cities, villageNames)
	// 设置工作簿的默认工作表
	f.SetActiveSheet(0)
	// 修改sheet名
	newName := time.Now().Format("01.02")
	f.SetSheetName(sheetName, newName)
	println(newName)
	newPath := "/SelfAndBig_" + time.Now().Format("20060102") + ".xlsx"
	newPath = utils.GetCurrentAbPath() + newPath
	log.Println(newPath)
	println(newPath)
	if err := f.SaveAs(newPath); err != nil {
		fmt.Println(err)
	}
}

// WriteHead 先绘画出excel表格的顶部
func WriteHead(f *excelize.File, cities []model.ProvinceCity, sheetName string) {
	cityLen := len(cities)
	asciiC := []rune("C")[0] //C的ascii码
	//第一行
	f.SetCellStr(sheetName, "A1", "骆湖镇网格化排查情况数据汇总表")
	firstCol := getColNumber(cityLen, asciiC, 10, 1)
	f.MergeCell(sheetName, "A1", firstCol)
	//第二行
	dateStr := time.Now().Format("2006.01.02")
	f.SetCellStr(sheetName, "A2", "填报日期："+dateStr)
	f.MergeCell(sheetName, "A2", getColNumber(cityLen, asciiC, 10, 2))
	//第三行
	f.SetCellStr(sheetName, "A3", "序号")
	f.MergeCell(sheetName, "A3", "A5")
	f.SetCellStr(sheetName, "B3", "村别")
	f.MergeCell(sheetName, "B3", "B5")
	f.SetCellStr(sheetName, "C3", "排查汇总")
	f.MergeCell(sheetName, "C3", "C5")
	f.SetCellStr(sheetName, "D3", "省市")
	f.MergeCell(sheetName, "D3", getColNumber(cityLen, asciiC, 0, 4))
	//
	//推送数
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 1, 3), "推送数")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 1, 3), getColNumber(cityLen, asciiC, 1, 5))
	//已排查
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 2, 3), "已排查")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 2, 3), getColNumber(cityLen, asciiC, 6, 3))
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 7, 3), "移交")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 7, 3), getColNumber(cityLen, asciiC, 7, 5))
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 8, 3), "重复")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 8, 3), getColNumber(cityLen, asciiC, 8, 5))
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 9, 3), "未排查")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 9, 3), getColNumber(cityLen, asciiC, 9, 5))
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 10, 3), "备注")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 10, 3), getColNumber(cityLen, asciiC, 10, 5))
	//第四行 落地排查
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 2, 4), "落地排查")
	f.MergeCell(sheetName, getColNumber(cityLen, asciiC, 2, 4), getColNumber(cityLen, asciiC, 6, 4))
	//第五行
	//城市列表
	for i, v := range cities {
		f.SetCellStr(sheetName, getColNumber(i+1, asciiC, 0, 5), v.CityName)
	}

	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 2, 5), "集中隔离")
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 3, 5), "居家隔离")
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 4, 5), "居家监测")
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 5, 5), "'四个一'")
	f.SetCellStr(sheetName, getColNumber(cityLen, asciiC, 6, 5), "3天居家监测+\n11天自我健康监测，三天两检")

}

// WriteBodyLeft 绘画左侧内容
func WriteBodyLeft(f *excelize.File, cities []model.ProvinceCity, sheetName string) []string {
	villages := GetVillageCompany(cities)
	asciiC := []rune("C")[0] //C的ascii码
	for i, v := range villages {
		//序号
		f.SetCellInt(sheetName, "A"+strconv.Itoa(i+6), i+1)
		//村名
		f.SetCellStr(sheetName, "B"+strconv.Itoa(i+6), v)
		//遍历cities,看看该村有无数据
		for j, city := range cities {
			for _, vv := range city.Villages {
				if v == vv.Name {
					//城市对应村居的数值
					colNum := getColNumber(j+1, asciiC, 0, i+6)
					f.SetCellInt(sheetName, colNum, vv.Number)
					break
				}
			}
		}
		//C列汇总，设置公式,第i行为该行所有城市对应村居数值的和
		formula := "=SUM(D" + strconv.Itoa(i+6) + ":" + getColNumber(len(cities), asciiC, 0, i+6) + ")"
		err := f.SetCellFormula(sheetName, "C"+strconv.Itoa(i+6), formula)
		if err != nil {
			log.Println(fmt.Sprintf("设置公式失败,错误:%s", err))
			fmt.Println(fmt.Sprintf("设置公式失败,错误:%s", err))
			return nil
		}
	}
	return villages
}

// WriteBodyRight 绘画右侧大数据表内容
func WriteBodyRight(f *excelize.File, f1SheetName string, path string, cities []model.ProvinceCity, villageNames []string) error {
	asciiC := []rune("C")[0] //C的ascii码
	cityLen := len(cities)
	bigData, err := ReadBigDataExcel(path)
	if err != nil {
		return err
	}

	//先填充村庄数据
	for i, _ := range bigData.Villages {
		for j, v := range bigData.Villages[i] {
			colNum := getColNumber(cityLen, asciiC, j+2, i+6)
			if v != 0 {
				f.SetCellInt(f1SheetName, colNum, v)
			}
		}
	}
	//填充企业数据,如果有相同企业，则大数据那边数据直接在这行添加
	for _, name := range bigData.EnterpriseName {
		for j, vn := range villageNames {
			//如果有相同企业，则大数据那边数据直接在这行添加
			if vn == name {
				enterprise := bigData.EnterpriseMap[name]
				array := enterprise.Arr
				//添加该企业一行数据
				for i, a := range array {
					colNum := getColNumber(cityLen, asciiC, i+2, j+6)
					if a != 0 {
						f.SetCellInt(f1SheetName, colNum, a)
					}
				}
				//添加完后，标志设为1
				enterprise.Flag = 1
				bigData.EnterpriseMap[name] = enterprise
				break
			}
		}
	}
	//填充企业数据，无相同企业，另起一行
	for _, name := range bigData.EnterpriseName {
		if bigData.EnterpriseMap[name].Flag == 0 {
			villageNames = append(villageNames, name)
			enterprise := bigData.EnterpriseMap[name]
			array := enterprise.Arr
			//添加该企业一行数据
			//序号
			f.SetCellStr(f1SheetName, "A"+strconv.Itoa(len(villageNames)+5), strconv.Itoa(len(villageNames)))
			//名字
			f.SetCellStr(f1SheetName, "B"+strconv.Itoa(len(villageNames)+5), name)
			for i, a := range array {
				colNum := getColNumber(cityLen, asciiC, i+2, len(villageNames)+5)
				if a != 0 {
					f.SetCellInt(f1SheetName, colNum, a)
				}
			}
		}
	}
	// 左侧村居企业已经最终确定
	villageLen := len(villageNames)
	//推送数
	pushNumber, err := strconv.Atoi(bigData.PushNumber)
	if err != nil {
		log.Println(err)
	}
	f.SetCellInt(f1SheetName, getColNumber(cityLen, asciiC, 1, 6), pushNumber)
	f.MergeCell(f1SheetName, getColNumber(cityLen, asciiC, 1, 6),
		getColNumber(cityLen, asciiC, 1, villageLen+5))
	// 合并右侧文本
	//转移
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 7, 6),
		utils.GetFormulaString(bigData.Transfer.Describe))
	f.MergeCell(f1SheetName, getColNumber(cityLen, asciiC, 7, 6),
		getColNumber(cityLen, asciiC, 7, villageLen+5))
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 7, villageLen+6), bigData.Transfer.Number)
	//重复
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 8, 6),
		utils.GetFormulaString(bigData.Repeat.Describe))
	f.MergeCell(f1SheetName, getColNumber(cityLen, asciiC, 8, 6),
		getColNumber(cityLen, asciiC, 8, villageLen+5))
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 8, villageLen+6), bigData.Repeat.Number)
	//未排查
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 9, 6),
		utils.GetFormulaString(bigData.NeverCheck.Describe))
	f.MergeCell(f1SheetName, getColNumber(cityLen, asciiC, 9, 6),
		getColNumber(cityLen, asciiC, 9, villageLen+5))
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 9, villageLen+6), bigData.NeverCheck.Number)
	//备注
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 10, 6),
		utils.GetFormulaString(bigData.Remark.Describe))
	f.MergeCell(f1SheetName, getColNumber(cityLen, asciiC, 10, 6),
		getColNumber(cityLen, asciiC, 10, villageLen+5))
	f.SetCellStr(f1SheetName, getColNumber(cityLen, asciiC, 10, villageLen+6), bigData.Remark.Number)

	// 小计和最后一行汇总
	f.SetCellStr(f1SheetName, "A"+strconv.Itoa(villageLen+6), "小计")
	f.MergeCell(f1SheetName, "A"+strconv.Itoa(villageLen+6), "B"+strconv.Itoa(villageLen+6))
	f.SetCellStr(f1SheetName, "A"+strconv.Itoa(villageLen+7), "汇总")
	f.MergeCell(f1SheetName, "A"+strconv.Itoa(villageLen+7), "B"+strconv.Itoa(villageLen+7))

	// 小计 公式
	calNum := cityLen + 7
	for i := 0; i < calNum; i++ {
		colnum1 := getColNumber(0, asciiC, i, 6)
		colnum2 := getColNumber(0, asciiC, i, villageLen+5)
		colnum3 := getColNumber(0, asciiC, i, villageLen+6)
		f.SetCellFormula(f1SheetName, colnum3, "=SUM("+colnum1+":"+colnum2+")")
	}
	// 汇总
	colnumA := getColNumber(0, asciiC, 0, villageLen+6)
	colnumB := getColNumber(cityLen, asciiC, 1, villageLen+6)
	f.SetCellFormula(f1SheetName, "C"+strconv.Itoa(villageLen+7),
		"=SUM("+colnumA+"+"+colnumB+")")
	f.MergeCell(f1SheetName, "C"+strconv.Itoa(villageLen+7),
		getColNumber(cityLen, asciiC, 10, villageLen+7))

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
	f.SetCellStyle(f1SheetName, "A1", getColNumber(cityLen, asciiC, 10, villageLen+7), styleId)
	return nil
}

// ReadBigDataExcel 读取大数据excel表数据
func ReadBigDataExcel(path string) (*model.BigData, error) {
	// 先读大数据表数据
	f2, err := excelize.OpenFile(path)
	if err != nil {
		log.Println(err)
		fmt.Println(err)
		return nil, err
	}
	// 获取 Sheet1 上所有单元格
	f2SheetName := f2.GetSheetName(0)
	f2rows, err := f2.GetRows(f2SheetName)
	if err != nil {
		log.Println(err)
		fmt.Println(err)
		return nil, err
	}
	//主要数据存储
	bigData := &model.BigData{}
	bigData.EnterpriseMap = map[string]model.EnterpriseArray{}
	for rowIndex, row := range f2rows {
		if rowIndex < 5 {
			continue
		}
		if rowIndex == 5 {
			for colIndex, colCell := range row {
				//推送数
				if colIndex == 2 {
					bigData.PushNumber = colCell
				}
				//移交
				if colIndex == 8 {
					info := model.DataInfo{}
					info.Describe = utils.GetFormulaString(colCell)
					bigData.Transfer = info
				}
				//重复
				if colIndex == 9 {
					info := model.DataInfo{}
					info.Describe = utils.GetFormulaString(colCell)
					bigData.Repeat = info
				}
				//未排查
				if colIndex == 10 {
					info := model.DataInfo{}
					info.Describe = utils.GetFormulaString(colCell)
					bigData.NeverCheck = info
				}
				//备注
				if colIndex == 11 {
					info := model.DataInfo{}
					info.Describe = utils.GetFormulaString(colCell)
					bigData.Remark = info
				}
			}
		}
		if rowIndex >= 6 && rowIndex <= 15 { // 获取数据
			for colIndex, colCell := range row {
				if colIndex >= 3 && colIndex <= 7 {
					if colCell == "" {
						continue
					}
					colint, err := strconv.Atoi(colCell)
					if err != nil {
						log.Println(err)
						println(err)
					}
					bigData.Villages[rowIndex-6][colIndex-3] = colint
				}
			}
		} else if rowIndex == len(f2rows)-1 { //最后一行，小计
			for colIndex, colCell := range row {
				switch colIndex {
				case 8:
					bigData.Transfer.Number = colCell
				case 9:
					bigData.Repeat.Number = colCell
				case 10:
					bigData.NeverCheck.Number = colCell
				case 11:
					bigData.Remark.Number = colCell
				}
			}
		} else if rowIndex > 15 { //企业,先记录，记录企业名数组和对应的hash数组
			//企业常量数组
			emps := constant.Villages[10:]
			curName := ""
			ea := model.EnterpriseArray{}
			for colIndex, colCell := range row {

				if colIndex == 1 {
					for _, vi := range emps {
						if strings.Contains(colCell, vi) {
							curName = vi
							break
						}
					}
					//未出现过的企业名，直接添加单元格名字作为企业名
					if curName == "" {
						curName = colCell
					}
					bigData.EnterpriseName = append(bigData.EnterpriseName, curName)
				} else if colIndex >= 3 && colIndex <= 7 {
					if colCell != "" {
						ea.Arr[colIndex-3], err = strconv.Atoi(colCell)
						log.Println(err)
						println(err)
					} else {
						ea.Arr[colIndex-3] = 0
					}
					//结束
					if colIndex == 7 {
						bigData.EnterpriseMap[curName] = ea
						break
					}
				}
			}
		}
	}
	return bigData, nil
}

// GetVillageCompany 得到村居和企业名字的数组
func GetVillageCompany(cities []model.ProvinceCity) []string {
	m := map[string]int{}
	villages := constant.Villages[0:10]
	for i, v := range villages {
		m[v] = i
	}
	for _, city := range cities {
		for _, v := range city.Villages {
			if _, ok := m[v.Name]; !ok {
				villages = append(villages, v.Name)
				m[v.Name] = len(villages) + 1
			}
		}
	}
	return villages
}

// 函数,给定动态列数值num,和固定列号stableCol(动态列的前一列所在列号,如动态列从D开始，那么此值为C),
//和忽略动态列的所求列的偏移数shift，以及行号，换取excel列号
func getColNumber(num int, stableCol rune, shift int, rows int) string {
	index := num + shift
	col := stableCol + rune(index)
	return string(col) + strconv.Itoa(rows)
}
