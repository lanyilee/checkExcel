package utils

import (
	"baliance.com/gooxml/document"
	"fmt"
	"io/fs"
	"io/ioutil"
	"log"
	"strconv"
	"strings"
)

// FileForEach 只读取当前指定路径下的文件，跳过文件夹
func FileForEach(fileFullPath string) ([]fs.FileInfo, error) {
	files, err := ioutil.ReadDir(fileFullPath)
	if err != nil {
		log.Fatal(err)
		return nil, err
	}
	var myFile []fs.FileInfo
	for _, file := range files {
		fmt.Println(file.Name())
		if file.IsDir() {
			continue
		}
		myFile = append(myFile, file)
	}
	return myFile, nil
}

// FileForEachComplete 读取指定路径下的所有文件，包括子文件夹下的
func FileForEachComplete(fileFullPath string) []fs.FileInfo {
	files, err := ioutil.ReadDir(fileFullPath)
	if err != nil {
		log.Fatal(err)
	}
	var myFile []fs.FileInfo
	for _, file := range files {
		if file.IsDir() {
			path := strings.TrimSuffix(fileFullPath, "/") + "/" + file.Name()
			subFile := FileForEachComplete(path)
			if len(subFile) > 0 {
				myFile = append(myFile, subFile...)
			}
		} else {
			myFile = append(myFile, file)
		}
	}
	return myFile
}

// ReadWordParagraphs 读取word.docx 文档段落
func ReadWordParagraphs() {
	doc, err := document.Open("/Users/lanyi/Documents/zzb/diaodong/2/东组公介字【2021】175号（罗涛）.docx")
	//doc, err := document.Open("/Users/lanyi/Documents/zzb/dian.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}
	//doc.Paragraphs()得到包含文档所有的段落的切片
	for i, para := range doc.Paragraphs() {
		//run为每个段落相同格式的文字组成的片段
		fmt.Println("-----------第", i, "段-------------")
		for j, run := range para.Runs() {
			fmt.Print("\t-----------第", j, "格式片段-------------")
			fmt.Print(run.Text())
		}
		fmt.Println()
	}
}

// ReadWordTable 读取word.docx 表格
func ReadWordTable() {
	doc, err := document.Open("/Users/lanyi/Documents/zzb/diaodong/2/东组公介字【2021】175号（罗涛）.docx")
	//doc, err := document.Open("/Users/lanyi/Documents/zzb/dian.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}
	//得到包含文档所有表格
	for tabId, tbl := range doc.Tables() { //返回文档类所有表格
		for rowId, row := range tbl.Rows() {
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
				fmt.Println("table"+strconv.Itoa(tabId), "行"+strconv.Itoa(rowId), "列"+strconv.Itoa(cellId), tex)
			}
		}
	}
}
