package vt

import (
	"github.com/ying32/govcl/vcl"
	"github.com/ying32/govcl/vcl/types"
)

type TMainForm struct {
	*vcl.TForm
}

var (
	MainForm *TMainForm
)

func (f *TMainForm) OnFormCreate(sender vcl.IObject) {
	f.SetCaption("导表管理器")
	f.SetBorderStyle(types.BsSingle)
	f.EnabledMaximize(false)
	f.SetWidth(960)
	f.SetHeight(600)
	f.ScreenCenter()
	f.newLeftGpBox()
	f.newRightGpBox()
}
