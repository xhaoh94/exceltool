package vt

import (
	"exceltool/cfg"
	"strings"
	"unsafe"

	"github.com/ying32/govcl/vcl"
	"github.com/ying32/govcl/vcl/types"
)

var (
	edit            *vcl.TEdit
	btnFind         *vcl.TButton
	chkListBox      *vcl.TCheckListBox
	btnAllSelect    *vcl.TButton
	btnCancelSelect *vcl.TButton
)

var (
	chkData []*cfg.CheckBoxStruct = make([]*cfg.CheckBoxStruct, 0)
)

func (f *TMainForm) newLeftGpBox() {
	pnl := vcl.NewFrame(f)
	pnl.SetParent(f)
	pnl.SetWidth(435)
	pnl.SetAlign(types.AlLeft)

	gpBox := vcl.NewGroupBox(pnl)
	gpBox.SetParent(pnl)
	gpBox.SetCaption("表格列表")
	gpBox.SetBounds(10, 10, 425, 580)

	edit = vcl.NewEdit(gpBox)
	edit.SetAutoSelect(false)
	edit.SetParent(gpBox)
	edit.SetLeft(10)
	edit.SetTop(0)
	edit.SetWidth(310)
	edit.SetTextHint("请输入名字查询")

	btnFind = vcl.NewButton(gpBox)
	btnFind.SetParent(gpBox)
	btnFind.SetCaption("搜索")
	btnFind.SetBounds(330, 0, 80, 25)
	btnFind.SetOnClick(onBtnFindClick)

	chkListBox = vcl.NewCheckListBox(gpBox)
	chkListBox.SetParent(gpBox)
	chkListBox.SetBounds(10, 30, 400, 480)

	btnAllSelect = vcl.NewButton(gpBox)
	btnAllSelect.SetParent(gpBox)
	btnAllSelect.SetCaption("全选")
	btnAllSelect.SetBounds(10, 520, 80, 30)
	btnAllSelect.SetOnClick(onBtnAllSelect)

	btnCancelSelect = vcl.NewButton(gpBox)
	btnCancelSelect.SetParent(gpBox)
	btnCancelSelect.SetCaption("取消选择")
	btnCancelSelect.SetBounds(100, 520, 80, 30)
	btnCancelSelect.SetOnClick(onBtnCancelSelect)

	updChkData()
}
func onBtnAllSelect(sender vcl.IObject) {
	chkListBox.CheckAll(types.CbChecked, true, true)
}
func onBtnCancelSelect(sender vcl.IObject) {
	chkListBox.CheckAll(types.CbUnchecked, true, true)
}
func onBtnFindClick(sender vcl.IObject) {
	updChkListBox()
}
func updChkListBox() {
	chkListBox.Clear()
	if len(chkData) == 0 {
		return
	}
	name := edit.Text()

	tmpList := make([]*cfg.CheckBoxStruct, 0)
	for _, v := range chkData {
		if find := strings.Contains(v.Name, name); name == "" || find {
			tmpList = append(tmpList, v)
		}
	}

	for _, v := range tmpList {
		chkListBox.Items().AddObject(v.Name, vcl.AsObject(unsafe.Pointer(v)))
	}
}

func updChkData() {
	path := cfg.GetInPath()
	if path == "" {
		return
	}
	chkData = cfg.ReadXlsx(path)
	updChkListBox()
}
