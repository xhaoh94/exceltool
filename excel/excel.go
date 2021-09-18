package excel

import (
	"encoding/json"
	"exceltool/cfg"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
	"github.com/ying32/govcl/vcl"
)

func ExportXlsx(cfgs []*cfg.OutCfg, list []*cfg.CheckBoxStruct) {
	fmt.Println("开始执行导出")
	fmt.Println("")
	exportPaths := make([]string, 0)
	exportDatas := make([][]byte, 0)
	for _, v := range list {
		if xlFile, err := xlsx.OpenFile(v.Path); err != nil {
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			fmt.Printf("打开xlxs文件出错 错误信息:[%s] 路径:[%s]\n", err, v.Path)
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			return
		} else {
			if paths, bs, stop := analysisFile(xlFile, cfgs); stop {
				return
			} else {
				exportPaths = append(exportPaths, paths...)
				exportDatas = append(exportDatas, bs...)
			}
		}
	}
	maxLen := len(exportPaths)
	for i := 0; i < maxLen; i++ {
		expath := exportPaths[i]
		b := exportDatas[i]
		dir := path.Dir(expath)
		err := os.MkdirAll(dir, 777)
		if err != nil {
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			fmt.Printf("创建文件夹时错误，写入路径:[%s]\n", expath)
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			return
		}
		err = ioutil.WriteFile(expath, b, 0666)
		if err != nil {
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			fmt.Printf("写入文件时错误，写入路径:[%s]\n写入数据:[%s]\n", expath, b)
			fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
			return
		}
	}
	fmt.Println("执行导出完成")
	vcl.ShowMessage("导出完成")
}

func analysisFile(xlFile *xlsx.File, cfgs []*cfg.OutCfg) ([]string, [][]byte, bool) {
	exportPaths := make([]string, 0)
	exportDatas := make([][]byte, 0)
	for _, sheet := range xlFile.Sheets {
		if path, b, stop := analysisSheet(sheet, cfgs); stop {
			return nil, nil, true
		} else {
			exportPaths = append(exportPaths, path...)
			exportDatas = append(exportDatas, b...)
		}
	}
	return exportPaths, exportDatas, false
}

func analysisSheet(sheet *xlsx.Sheet, cfgs []*cfg.OutCfg) ([]string, [][]byte, bool) {
	sheetStruct := getCelValue(sheet, 0, 1)
	if sheetStruct == "" {
		return nil, nil, false
	}
	exportPaths := make([]string, 0)
	exportDatas := make([][]byte, 0)

	exportPath := getCelValue(sheet, 1, 1)
	keyCnt := getKeyCnt(sheet, 2, 1)

	cols, fields := getFields(sheet)
	tags := getTags(sheet)
	for i := range cfgs {
		cfg := cfgs[i]
		values := getValues(sheet, cfg.TagType, cols, fields, tags)
		if len(values) == 0 {
			continue
		}
		keyFields := make([]string, 0)
		temField2Vs := make([]string, 0)
		for key := 1; key <= keyCnt; key++ {
			keyField := fields[key]
			keyFields = append(keyFields, keyField)
			temField2Vs = append(temField2Vs, "")
		}
		str := ""
		first := true
		for i, jsonStr := range values {
			var f interface{}
			json.Unmarshal([]byte(jsonStr), &f)
			m := f.(map[string]interface{})
			for keyIndex, keyField := range keyFields {
				keyV := strval(m[keyField])
				if keyV == "" {
					fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
					fmt.Printf("使用了非法类型作为主KEY,表名[%s]\n", sheet.Name)
					fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
					return nil, nil, true
				}
				if temField2Vs[keyIndex] != keyV {
					oldVs := temField2Vs[keyIndex]
					temField2Vs[keyIndex] = keyV
					//重置临时数据key
					for key := keyIndex + 1; key < keyCnt; key++ {
						temField2Vs[key] = ""
					}

					if keyIndex == keyCnt-1 {
						if !first {
							str += ","
						}
						first = false
						str += "\"" + strval(keyV) + "\":" + jsonStr
					} else {
						if keyIndex == keyCnt-2 {
							first = true
						}
						if oldVs != "" {
							for l := keyIndex + 1; l < keyCnt; l++ {
								str += "}"
							}
							str += ","
						}
						str += "\"" + strval(keyV) + "\":" + "{"
					}
				} else {
					if keyIndex == keyCnt-1 {
						fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
						fmt.Printf("出现相同KEY值,表名[%s],行数[%d]\n", sheet.Name, i+8)
						fmt.Println("⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐")
						return nil, nil, true
					}
				}
			}
		}
		for l := 1; l < keyCnt; l++ {
			str += "}"
		}
		str = "{\"" + sheetStruct + "\":{" + str + "}}"
		path := path.Join(cfg.OutPath, exportPath)

		exportPaths = append(exportPaths, path)
		exportDatas = append(exportDatas, []byte(str))
	}

	fmt.Println("///////////////////////////正在解析/////////////////////////////////")
	fmt.Println("表名: ", sheet.Name)
	fmt.Println("导出结构体: ", sheetStruct)
	fmt.Println("key数量: ", keyCnt)
	fmt.Println("字段集合: ", fields)
	fmt.Println("标签集合: ", tags)
	for i := range exportPaths {
		fmt.Println("导出文件: ", exportPaths[i])
	}
	fmt.Println("///////////////////////////解析完成//////////////////////////////////")
	fmt.Println("")

	return exportPaths, exportDatas, false
}

func getCelValue(sheet *xlsx.Sheet, row int, col int) string {
	if c, err := sheet.Cell(row, col); err != nil {
		return ""
	} else {
		return c.Value
	}
}
func getKeyCnt(sheet *xlsx.Sheet, row int, col int) int {
	if c, err := sheet.Cell(row, col); err != nil {
		return 0
	} else {
		if key, err := c.Int(); err != nil {
			return 0
		} else {
			return key
		}
	}
}
func getTags(sheet *xlsx.Sheet) map[int]string {
	tags := make(map[int]string, 0)
	if row, err := sheet.Row(5); err == nil {
		err := row.ForEachCell(func(c *xlsx.Cell) error {
			if colNum, _ := c.GetCoordinates(); colNum > 0 {
				tags[colNum] = c.Value
			}
			return nil
		}, xlsx.SkipEmptyCells)
		if err != nil {
			return nil
		}
	}
	return tags
}
func getFields(sheet *xlsx.Sheet) ([]int, map[int]string) {
	fields := make(map[int]string, 0)
	if row, err := sheet.Row(6); err == nil {
		err := row.ForEachCell(func(c *xlsx.Cell) error {
			if colNum, _ := c.GetCoordinates(); colNum > 0 {
				fields[colNum] = c.Value
			}
			return nil
		}, xlsx.SkipEmptyCells)
		if err != nil {
			return nil, nil
		}
	}
	cols := make([]int, 0)
	for key := range fields {
		cols = append(cols, key)
	}
	sort.Ints(cols)

	return cols, fields
}

func getValues(sheet *xlsx.Sheet, tagA string, cols []int, fields map[int]string, tags map[int]string) []string {
	values := make([]string, 0)
	for i := 7; i < sheet.MaxRow; i++ {
		jsonStr := "{"
		for index, col := range cols {
			tagB := tags[col]
			if tagB == "" {
				continue
			}
			field := fields[col]
			if field == "" {
				continue
			}
			if !checkTag(tagA, tagB) {
				continue
			}
			if c, err := sheet.Cell(i, col); err != nil {
				continue
			} else {
				if c.Value != "" {
					if index != 0 {
						jsonStr += ","
					}
					jsonStr += "\"" + field + "\":" + c.Value
				}
			}
		}
		jsonStr += "}"
		if jsonStr == "{}" {
			break
		}
		values = append(values, jsonStr)
	}
	return values
}
func checkTag(tagA string, tagB string) bool {
	if find := strings.Contains(tagB, tagA); find {
		return true
	}
	return false
}

// strval 获取变量的字符串值
// 浮点型 3.0将会转换成字符串3, "3"
// 非数值或字符类型的变量将会被转换成JSON格式字符串
func strval(value interface{}) string {
	var key string
	if value == nil {
		return key
	}

	switch value.(type) {
	case float64:
		ft := value.(float64)
		key = strconv.FormatFloat(ft, 'f', -1, 64)
	case float32:
		ft := value.(float32)
		key = strconv.FormatFloat(float64(ft), 'f', -1, 64)
	case int:
		it := value.(int)
		key = strconv.Itoa(it)
	case uint:
		it := value.(uint)
		key = strconv.Itoa(int(it))
	case int8:
		it := value.(int8)
		key = strconv.Itoa(int(it))
	case uint8:
		it := value.(uint8)
		key = strconv.Itoa(int(it))
	case int16:
		it := value.(int16)
		key = strconv.Itoa(int(it))
	case uint16:
		it := value.(uint16)
		key = strconv.Itoa(int(it))
	case int32:
		it := value.(int32)
		key = strconv.Itoa(int(it))
	case uint32:
		it := value.(uint32)
		key = strconv.Itoa(int(it))
	case int64:
		it := value.(int64)
		key = strconv.FormatInt(it, 10)
	case uint64:
		it := value.(uint64)
		key = strconv.FormatUint(it, 10)
	case string:
		key = value.(string)
	case []byte:
		key = string(value.([]byte))
	default:
		key = ""
	}

	return key
}
