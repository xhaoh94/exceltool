package main

import (
	"exceltool/cfg"
	"exceltool/vt"

	_ "github.com/ying32/govcl/pkgs/winappres"
	"github.com/ying32/govcl/vcl"
)

func main() {
	if cfg.ReadCfg() {
		runApp()
	}

	// files := cfg.ReadXlsx("D:/Project/Golang/exceltool/sheets")
	// excel.ReadXlsx(files[0].Path, "s")
}

func runApp() {

	vcl.Application.Initialize()
	vcl.Application.SetMainFormOnTaskBar(true)
	vcl.Application.CreateForm(&vt.MainForm)
	vcl.Application.CreateForm(&vt.OutForm)
	vcl.Application.Run()

	// vcl.RunApp(&vt.MainForm, &vt.OutForm)
}
