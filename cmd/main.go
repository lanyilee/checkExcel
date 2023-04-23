package main

import (
	"fmt"
	"fyno/app"
	"fyno/utils"
	"github.com/flopp/go-findfont"
	"log"
	"os"
	"strings"
	"time"
)

// 方式一  设置环境变量   通过go-findfont 寻找simkai.ttf 字体
func init() {
	fontPaths := findfont.List()
	for _, fontPath := range fontPaths {
		fmt.Println(fontPath)
		//楷体:simkai.ttf
		//黑体:simhei.ttf
		if strings.Contains(fontPath, "simkai.ttf") {
			err := os.Setenv("FYNE_FONT", fontPath)
			if err != nil {
				return
			}
			break
		}
	}

	InitLog()
}

func main() {
	//app.NewCanvas()
	//app.MainShow()

	app.NewTransferEntry()

}

func InitLog() {
	//日志输出位置配置
	//log.SetFlags(log.Ldate | log.Ltime | log.Lshortfile)
	log.SetFlags(log.Ltime | log.Lshortfile)
	date := time.Now().Format("2006-01-02")
	dateLog := utils.GetCurrentAbPath() + "/" + date + ".log"
	println("log path: " + dateLog)
	logFile, err := os.OpenFile(dateLog, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		log.Panic("打开日志文件异常")
	}
	log.SetOutput(logFile)
}
