package model

type Unit struct {
	UnitName string
	// 较好
	FirstAboutGood int
	// 好
	FirstGood int

	SecondAboutGood int
	SecondGood      int

	ThirdAboutGood int
	ThirdGood      int

	FourthAboutGood int
	FourthGood      int

	IfPrint bool
}

type UnitPerson struct {
	UnitName            string
	FirstQuarter        string
	SecondQuarter       string
	ThirdQuarter        string
	FourthQuarter       string
	FirstQuarterRecord  string
	SecondQuarterRecord string
	ThirdQuarterRecord  string
	FourthQuarterRecord string
	Type                string
}
