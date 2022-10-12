package model


type Person struct {
	Number string
	Name string
	Sex string
	Age string
	Id string
	Phone string
	Village string
	LeaveTime string
	BackTime string
	Control string
	ControlTime string
	PrevAddr string
}

type ProvinceCity struct {
	CityName string
	Index int //在常量表中的下标
	Number int
	Villages []Village
}

type Village struct {
	Name string
	Number int
}