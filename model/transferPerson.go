package model

type TransferPerson struct {
	Name     string
	Out      string //调出单位
	In       string //调入单位
	WorkDate string //办理时间
	Remark   string
	IfNew    bool   //是否是新录入
	Session  string //zhuanrenqingk
}
