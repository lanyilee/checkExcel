package model

// BigData 大数据推送表格
type BigData struct {
	//村庄数据
	Villages [10][5]int
	//企业数据
	EnterpriseName []string
	EnterpriseMap map[string]EnterpriseArray
	//移交
	Transfer DataInfo
	//重复
	Repeat DataInfo
	//未排查
	NeverCheck DataInfo
	//备注
	Remark DataInfo
	//推送数
	PushNumber string

}

// DataInfo 移交等具体信息
type DataInfo struct {
	Describe string
	Number string
}

// EnterpriseArray 企业对应数据数组,flag 是该企业数据是否已被添加的标志,0为未添加
type EnterpriseArray struct {
	Arr [5]int
	Flag int
}
