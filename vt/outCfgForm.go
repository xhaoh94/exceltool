package vt

import (
	"exceltool/cfg"
	"path/filepath"

	"github.com/google/uuid"
	"github.com/ying32/govcl/vcl"
	"github.com/ying32/govcl/vcl/types"
)

type TAddOutCfgForm struct {
	*vcl.TForm
	edit            *vcl.TEdit
	pathLb          *vcl.TLabel
	btnSelectDirDia *vcl.TButton
	cb              *vcl.TCheckBox
	btnApply        *vcl.TButton
	exportCombox    *vcl.TComboBox
	tagCombox       *vcl.TComboBox
}

var (
	OutForm *TAddOutCfgForm
)

func (f *TAddOutCfgForm) OnFormCreate(sender vcl.IObject) {

	f.SetCaption("添加配置")
	f.SetBorderStyle(types.BsSingle)
	f.EnabledMaximize(false)
	f.SetWidth(300)
	f.SetHeight(250)
	f.SetBorderIcons(1)
	f.ScreenCenter()

	gpBox := vcl.NewGroupBox(f)
	gpBox.SetParent(f)
	gpBox.SetCaption("导出路径")
	gpBox.SetBounds(10, 10, 280, 60)

	gp := vcl.NewGroupBox(gpBox)
	gp.SetParent(gpBox)
	gp.SetBounds(10, 0, 220, 30)

	f.pathLb = vcl.NewLabel(gp)
	f.pathLb.SetParent(gp)
	f.pathLb.SetBounds(0, -11, 200, 20)

	f.btnSelectDirDia = vcl.NewButton(gpBox)
	f.btnSelectDirDia.SetParent(gpBox)
	f.btnSelectDirDia.SetCaption("···")
	f.btnSelectDirDia.SetBounds(230, 6, 30, 23)
	f.btnSelectDirDia.SetOnClick(f.onBtnSelectDirDiaClick)

	lb := vcl.NewLabel(f)
	lb.SetParent(f)
	lb.SetCaption("导出类型:")
	lb.SetTop(83)
	lb.SetLeft(10)
	// TComboBox
	f.exportCombox = vcl.NewComboBox(f)
	f.exportCombox.SetParent(f)
	f.exportCombox.SetTop(80)
	f.exportCombox.SetLeft(70)
	f.exportCombox.SetStyle(types.CsDropDownList)
	for _, v := range cfg.ExportType {
		f.exportCombox.Items().Add(v)
	}
	f.exportCombox.SetItemIndex(0)

	lb = vcl.NewLabel(f)
	lb.SetParent(f)
	lb.SetCaption("导出标签:")
	lb.SetTop(113)
	lb.SetLeft(10)
	// TComboBox
	f.tagCombox = vcl.NewComboBox(f)
	f.tagCombox.SetParent(f)
	f.tagCombox.SetTop(110)
	f.tagCombox.SetLeft(70)
	f.tagCombox.SetStyle(types.CsDropDownList)
	for _, v := range cfg.TagType {
		f.tagCombox.Items().Add(v)
	}
	f.tagCombox.SetItemIndex(0)

	lbName := vcl.NewLabel(f)
	lbName.SetParent(f)
	lbName.SetCaption("配置名称:")
	lbName.SetTop(143)
	lbName.SetLeft(10)
	f.edit = vcl.NewEdit(f)
	f.edit.SetParent(f)
	f.edit.SetTextHint("请输入名字")
	f.edit.SetBounds(70, 140, 160, 20)

	f.btnApply = vcl.NewButton(f)
	f.btnApply.SetParent(f)
	f.btnApply.SetCaption("添加")
	f.btnApply.SetBounds(100, 180, 100, 50)
	f.btnApply.SetOnClick(f.onBtnApplyClick)

}
func (f *TAddOutCfgForm) onBtnSelectDirDiaClick(sender vcl.IObject) {
	sdd := vcl.NewSelectDirectoryDialog(nil)
	if sdd.Execute() {
		path := filepath.ToSlash(sdd.FileName())
		f.pathLb.SetCaption(path)
	}
}
func (f *TAddOutCfgForm) onBtnApplyClick(sender vcl.IObject) {
	path := f.pathLb.Caption()
	if path == "" {
		vcl.ShowMessage("导出路径不能为空")
		return
	}
	tn := f.edit.Text()
	if tn == "" {
		vcl.ShowMessage("配置名不能为空")
		return
	}

	exportStr := f.exportCombox.Items().Strings(f.exportCombox.ItemIndex())
	tagStr := f.tagCombox.Items().Strings(f.tagCombox.ItemIndex())
	outCfgs := cfg.GetOutPath()
	id := StringToHash(uuid.New().String())
	outCfg := &cfg.OutCfg{ID: id, Name: tn, OutPath: path, IsExport: true, ExportType: exportStr, TagType: tagStr}
	outCfgs = append(outCfgs, outCfg)
	cfg.SetOutPath(outCfgs)
	updOutPath()
	f.Close()
	MainForm.SetFocus()
}
