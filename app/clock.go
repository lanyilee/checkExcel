package app

import (
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/widget"
	time2 "time"
)

func NewClock()  {
	a := app.New()
	w := a.NewWindow("clock")

	clock := widget.NewLabel("")

	w.SetContent(clock)
	go func() {
		for range time2.Tick(time2.Second){
			updateClockTime(clock)
		}
	}()
	w.ShowAndRun()
}


func updateClockTime(clock *widget.Label){
	time := time2.Now().Format("Time: 03:04:05")
	clock.SetText(time)
}
