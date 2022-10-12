package app

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/widget"
)

func ThreeWindows(){
	a := app.New()
	w := a.NewWindow("Hello World")

	w.SetContent(widget.NewLabel("Hello World!"))
	w.Show()

	w2 := a.NewWindow("Larger")
	w2.SetContent(widget.NewButton("add window",func() {
		w3 := a.NewWindow("third")
		w3.SetContent(widget.NewLabel("third"))
		w3.Show()
	}))
	w2.Resize(fyne.NewSize(100, 100))
	w2.Show()

	a.Run()
}
