package vt

import (
	"exceltool/cfg"
	"exceltool/excel"
	"hash/crc32"
	"path/filepath"

	"github.com/ying32/govcl/vcl"
	"github.com/ying32/govcl/vcl/types"
)

var (
	btnSelectDirDia *vcl.TButton
	pathEdit        *vcl.TEdit
	btnAddCfg       *vcl.TButton
	btnDelCfg       *vcl.TButton
	pageCol         *vcl.TPageControl

	btnExportAll    *vcl.TButton
	btnExportSelect *vcl.TButton

	id2tab  map[int]*vcl.TTabSheet = make(map[int]*vcl.TTabSheet)
	tab2Cfg map[int]*cfg.OutCfg    = make(map[int]*cfg.OutCfg)
)

func (f *TMainForm) newPathBox(parent *vcl.TFrame) {

	pnl := vcl.NewFrame(parent)
	pnl.SetParent(parent)
	// pnl.SetColor(colors.ClYellow)
	pnl.SetHeight(80)
	pnl.SetAlign(types.AlTop)

	lb := vcl.NewLabel(pnl)
	lb.SetParent(pnl)
	lb.SetCaption("表格列表路径:")
	lb.SetBounds(10, 15, 505, 60)

	pathEdit = vcl.NewEdit(pnl)
	pathEdit.SetParent(pnl)
	pathEdit.SetEnabled(false)
	pathEdit.SetBounds(10, 35, 390, 20)
	updInPath()

	btnSelectDirDia = vcl.NewButton(pnl)
	btnSelectDirDia.SetParent(pnl)
	btnSelectDirDia.SetCaption("···")
	btnSelectDirDia.SetBounds(420, 35, 80, 25)
	btnSelectDirDia.SetOnClick(onBtnSelectDirDiaClick)
}
func (f *TMainForm) newCfgBox(parent *vcl.TFrame) {

	pnl := vcl.NewFrame(parent)
	pnl.SetParent(parent)
	pnl.SetAlign(types.AlClient)

	btnAddCfg = vcl.NewButton(pnl)
	btnAddCfg.SetParent(pnl)
	btnAddCfg.SetCaption("添加配置")
	btnAddCfg.SetBounds(10, 0, 100, 30)
	btnAddCfg.SetOnClick(onBtnAddCfgClick)

	btnDelCfg = vcl.NewButton(pnl)
	btnDelCfg.SetParent(pnl)
	btnDelCfg.SetCaption("删除配置")
	btnDelCfg.SetBounds(120, 0, 100, 30)
	btnDelCfg.SetOnClick(onBtnDelCfgClick)

	pageCol = vcl.NewPageControl(pnl)
	pageCol.SetParent(pnl)
	pageCol.SetBounds(10, 40, 505, 180)
	updOutPath()

}
func (f *TMainForm) newBtnBox(parent *vcl.TFrame) {

	pnl := vcl.NewFrame(parent)
	pnl.SetParent(parent)
	pnl.SetHeight(100)
	pnl.SetAlign(types.AlBottom)

	btnExportSelect = vcl.NewButton(pnl)
	btnExportSelect.SetParent(pnl)
	btnExportSelect.SetCaption("导出选中项")
	btnExportSelect.SetBounds(10, 25, 200, 50)
	btnExportSelect.SetOnClick(onBtnExportSelectClick)

	btnExportAll = vcl.NewButton(pnl)
	btnExportAll.SetParent(pnl)
	btnExportAll.SetCaption("导出全部")
	btnExportAll.SetBounds(220, 25, 200, 50)
	btnExportAll.SetOnClick(onBtnExportAllClick)
}
func (f *TMainForm) newRightGpBox() {
	pnl := vcl.NewFrame(f)
	pnl.SetParent(f)
	pnl.SetAlign(types.AlClient)
	f.newPathBox(pnl)
	f.newCfgBox(pnl)
	f.newBtnBox(pnl)
}

func onBtnSelectDirDiaClick(sender vcl.IObject) {
	sdd := vcl.NewSelectDirectoryDialog(nil)
	if sdd.Execute() {
		path := filepath.ToSlash(sdd.FileName())
		if cfg.SetInPath(path) {
			updInPath()
			updChkData()
		}
	}
}
func updInPath() {
	path := cfg.GetInPath()
	pathEdit.SetText(path)
}

func onBtnAddCfgClick(sender vcl.IObject) {
	OutForm.Show()
}
func onBtnDelCfgClick(sender vcl.IObject) {
	index := pageCol.ActivePageIndex()
	outCfg := cfg.GetOutPath()
	outCfg = append(outCfg[:index], outCfg[index+1:]...)
	cfg.SetOutPath(outCfg)
	updOutPath()
}

func onBtnExportAllClick(sender vcl.IObject) {
	onExport(chkData)
}
func onBtnExportSelectClick(sender vcl.IObject) {

	list := make([]*cfg.CheckBoxStruct, 0)
	maxLen := chkListBox.Items().Count()
	var i int32
	for i = 0; i < maxLen; i++ {
		if chkListBox.Checked(i) {
			obj := (*cfg.CheckBoxStruct)(chkListBox.Items().Objects(i).UnsafeAddr())
			list = append(list, obj)
		}
	}
	onExport(list)
}
func onExport(list []*cfg.CheckBoxStruct) {
	if len(list) == 0 {
		vcl.ShowMessage("请选中导出项")
		return
	}
	cfgs := make([]*cfg.OutCfg, 0)
	for key := range tab2Cfg {
		cfg := tab2Cfg[key]
		if !cfg.IsExport {
			continue
		}
		cfgs = append(cfgs, cfg)
	}
	excel.ExportXlsx(cfgs, list)
}

func updOutPath() {

	outCfgs := cfg.GetOutPath()
	temCfgs := make(map[int]bool)
	for _, v := range outCfgs {
		temCfgs[v.ID] = true
	}

	if len(id2tab) > 0 {
		for k := range id2tab {
			v := id2tab[k]
			if temCfgs[k] == false {
				delete(tab2Cfg, k)
				delete(id2tab, k)
				v.SetPageControl(nil)
			}
		}
	}
	for _, v := range outCfgs {
		tab := id2tab[v.ID]
		isNew := false
		if tab == nil {
			tab = vcl.NewTabSheet(pageCol)
			tab.SetTag(v.ID)
			tab.SetPageControl(pageCol)
			tab.SetCaption(v.Name)
			id2tab[v.ID] = tab
			tab2Cfg[v.ID] = v
			isNew = true
		}
		createOutCfgView(tab, isNew)
	}
}

func createOutCfgView(tab *vcl.TTabSheet, isNew bool) {

	tag := tab.Tag()
	outCfg := tab2Cfg[tag]
	var pathLb *vcl.TEdit
	var check *vcl.TCheckBox
	var exportCombox *vcl.TComboBox
	var tagCombox *vcl.TComboBox
	if isNew {

		title := vcl.NewStaticText(tab)
		title.SetParent(tab)
		title.SetCaption("导出路径:")
		title.SetBounds(10, 5, 200, 20)

		pathLb = vcl.NewEdit(tab)
		pathLb.SetName("pathLb")
		pathLb.SetParent(tab)
		pathLb.SetEnabled(false)
		pathLb.SetBounds(10, 30, 380, 20)

		btn := vcl.NewButton(tab)
		btn.SetParent(tab)
		btn.SetCaption("···")
		btn.SetBounds(410, 30, 80, 25)
		btn.SetOnClick(onTabSelectDirClick)

		exLb := vcl.NewLabel(tab)
		exLb.SetParent(tab)
		exLb.SetCaption("导出类型:")
		exLb.SetTop(65)
		exLb.SetLeft(10)

		exportCombox = vcl.NewComboBox(tab)
		exportCombox.SetName("exportCombox")
		exportCombox.SetParent(tab)
		exportCombox.SetTop(60)
		exportCombox.SetLeft(70)
		exportCombox.SetStyle(types.CsDropDownList)
		exportCombox.SetOnChange(onExportComboxChanged)
		for _, v := range cfg.ExportType {
			exportCombox.Items().Add(v)
		}

		exLb = vcl.NewLabel(tab)
		exLb.SetParent(tab)
		exLb.SetCaption("导出标签:")
		exLb.SetTop(95)
		exLb.SetLeft(10)
		// TComboBox
		tagCombox = vcl.NewComboBox(tab)
		tagCombox.SetName("tagCombox")
		tagCombox.SetParent(tab)
		tagCombox.SetTop(90)
		tagCombox.SetLeft(70)
		tagCombox.SetStyle(types.CsDropDownList)
		tagCombox.SetOnChange(onTagComboxChanged)
		for _, v := range cfg.TagType {
			tagCombox.Items().Add(v)
		}

		check = vcl.NewCheckBox(tab)
		check.SetName("check")
		check.SetParent(tab)
		check.SetTop(120)
		check.SetLeft(10)
		check.SetCaption("是否导出")
		check.SetOnChange(onTabCheckChanged)

	} else {
		pathLb = vcl.AsEdit(tab.FindChildControl("pathLb"))
		check = vcl.AsCheckBox(tab.FindChildControl("check"))
		exportCombox = vcl.AsComboBox(tab.FindChildControl("exportCombox"))
		tagCombox = vcl.AsComboBox(tab.FindChildControl("tagCombox"))
	}
	pathLb.SetText(outCfg.OutPath)
	check.SetChecked(outCfg.IsExport)
	index := 0
	for i, v := range cfg.ExportType {
		if outCfg.ExportType == v {
			index = i
		}
	}
	exportCombox.SetItemIndex(int32(index))
	index = 0
	for i, v := range cfg.TagType {
		if outCfg.TagType == v {
			index = i
		}
	}
	tagCombox.SetItemIndex(int32(index))
}

func onTabSelectDirClick(sender vcl.IObject) {
	sdd := vcl.NewSelectDirectoryDialog(nil)
	if sdd.Execute() {
		path := filepath.ToSlash(sdd.FileName())
		if path == "" {
			vcl.ShowMessage("导出路径不能为空")
			return
		}

		btn := vcl.AsButton(sender)
		tab := vcl.AsTabSheet(btn.Parent())
		tag := tab.Tag()
		// tab = id2tab[tag]
		// gp := vcl.AsGroupBox(tab.FindChildControl("gp"))
		pathLb := vcl.AsLabel(tab.FindChildControl("pathLb"))
		if pathLb.Caption() == path {
			return
		}
		pathLb.SetCaption(path)

		outCfg := tab2Cfg[tag]
		outCfg.OutPath = path
		cfg.WriteCfg()
	}
}
func onTabCheckChanged(sender vcl.IObject) {
	check := vcl.AsCheckBox(sender)
	tab := vcl.AsTabSheet(check.Parent())
	tag := tab.Tag()
	outCfg := tab2Cfg[tag]
	if outCfg.IsExport != check.Checked() {
		outCfg.IsExport = check.Checked()
		cfg.WriteCfg()
	}
}
func onExportComboxChanged(sender vcl.IObject) {
	combox := vcl.AsComboBox(sender)
	tab := vcl.AsTabSheet(combox.Parent())
	tag := tab.Tag()
	outCfg := tab2Cfg[tag]
	str := combox.Items().Strings(combox.ItemIndex())
	if outCfg.ExportType != str {
		outCfg.ExportType = str
		cfg.WriteCfg()
	}
}
func onTagComboxChanged(sender vcl.IObject) {
	combox := vcl.AsComboBox(sender)
	tab := vcl.AsTabSheet(combox.Parent())
	tag := tab.Tag()
	outCfg := tab2Cfg[tag]
	str := combox.Items().Strings(combox.ItemIndex())
	if outCfg.TagType != str {
		outCfg.TagType = str
		cfg.WriteCfg()
	}
}

//StringToHash 字符串转为32位整形哈希
func StringToHash(s string) (hash int) {

	hash = int(crc32.ChecksumIEEE([]byte(s)))
	if hash >= 0 {
		return hash
	}
	if -hash >= 0 {
		return -hash
	}

	for _, c := range s {
		ch := int(c)
		hash = hash + ((hash) << 5) + ch + (ch << 7)
	}
	return
}
