package cfg

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"strings"
)

type OutCfg struct {
	ID         int
	Name       string
	OutPath    string
	IsExport   bool
	ExportType string
	TagType    string
}
type ConfStruct struct {
	InPath   string
	OutPaths []*OutCfg
}

type CheckBoxStruct struct {
	Name     string
	Path     string
	FileInfo os.FileInfo
}

var (
	conf       *ConfStruct
	ExportType []string = []string{"Json"}
	TagType    []string
)

func init() {
	TagType = make([]string, 0)
	for i := 'a'; i <= 'z'; i++ {
		TagType = append(TagType, string(i))
	}
}

func initCfg() {
	conf = &ConfStruct{
		InPath:   "",
		OutPaths: make([]*OutCfg, 0),
	}
	WriteCfg()
}

func ReadCfg() bool {
	fileByte, fileErr := ioutil.ReadFile("config.json")
	if fileErr != nil {
		initCfg()
		return true
	}
	fileErr = json.Unmarshal(fileByte, &conf)
	if fileErr != nil {
		fmt.Println("读取config.json有误")
		return false
	}
	return true
}

func GetOutPath() []*OutCfg {
	return conf.OutPaths
}
func SetOutPath(v []*OutCfg) {
	conf.OutPaths = v
	WriteCfg()
}

func GetInPath() string {
	return conf.InPath
}

func SetInPath(v string) bool {
	if conf.InPath != v {
		conf.InPath = v
		WriteCfg()
		return true
	}
	return false
}

func WriteCfg() {
	bytes, _ := json.Marshal(&conf)
	ioutil.WriteFile("config.json", bytes, 0777)
}

func ReadXlsx(path string) []*CheckBoxStruct {
	results := make([]*CheckBoxStruct, 0)
	if filesInfo, err := ioutil.ReadDir(path); err != nil {
		err = errors.New("读取文件失败: " + path)
		return nil
	} else {
		var reg = regexp.MustCompile(`.\.{1}xlsx$`)
		for _, fileInfo := range filesInfo {
			if fileInfo.IsDir() {
				ReadXlsx(path + "/" + fileInfo.Name())
			} else {
				if reg.MatchString(fileInfo.Name()) && !strings.Contains(fileInfo.Name(), "~$") {
					results = append(results, &CheckBoxStruct{
						Name:     fileInfo.Name(),
						FileInfo: fileInfo,
						Path:     path + "/" + fileInfo.Name(),
					})
				}
			}
		}
	}
	return results
}
