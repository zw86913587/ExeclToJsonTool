package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

func main() {
	excelFiles, done := findExcelFiles()
	if done {
		return
	}

	// 遍历找到的每个 Excel 文件
	convertExcelToJSON(excelFiles)
}

func findExcelFiles() ([]string, bool) {
	// 获取当前目录
	currentDir, err := os.Getwd()
	if err != nil {
		log.Fatalf("无法获取当前目录: %v", err)
	}

	// 查找当前目录中的 Excel 文件
	var excelFiles []string
	err = filepath.Walk(currentDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		// 检查文件扩展名是否为 .xlsx 或 .xls
		if !info.IsDir() && (filepath.Ext(info.Name()) == ".xlsx" || filepath.Ext(info.Name()) == ".xls") && strings.Index(info.Name(), "~") == -1 {
			excelFiles = append(excelFiles, path)
		}
		return nil
	})

	if err != nil {
		log.Fatalf("遍历目录时出错: %v", err)
	}

	if len(excelFiles) == 0 {
		fmt.Println("当前目录中未找到任何 Excel 文件 (.xlsx 或 .xls)")
		return nil, true
	}
	return excelFiles, false
}

func convertExcelToJSON(excelFiles []string) {
	for _, excelFile := range excelFiles {
		jsonDir := filepath.Join(filepath.Dir(excelFile), "json")

		// 检查文件夹是否存在，如果不存在则创建它
		if err := os.MkdirAll(jsonDir, os.ModePerm); err != nil {
			log.Printf("无法创建文件夹 %s: %v", jsonDir, err)
			continue
		}

		// 生成输出文件名，去掉路径，替换扩展名
		fileName := filepath.Base(excelFile) // 获取文件名
		jsonFile := filepath.Join(jsonDir, fileName[:len(fileName)-len(filepath.Ext(fileName))]+".json")

		// 读取 Excel 文件
		f, err := excelize.OpenFile(excelFile)
		if err != nil {
			log.Printf("无法打开 Excel 文件 %s: %v", excelFile, err)
			continue
		}

		// 获取第一个工作表名称
		sheetName := f.GetSheetName(0)

		// 读取工作表数据
		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Printf("无法读取工作表 %s: %v", sheetName, err)
			continue
		}

		// 将数据转换为 JSON 格式
		data := make(map[string]map[string]string)
		if len(rows) > 0 {
			// 处理表头
			headers := rows[0]
			for _, row := range rows[4:] {
				entry := make(map[string]string)
				ikey := "-1"
				for i, cell := range row {
					if i+1 > len(rows[3]) {
						break // 跳过超出表头范围的行
					}
					if rows[3][i] != "3" {
						continue // 跳过不显示的列
					}
					if i == 0 {
						ikey = cell
					}
					// 尝试将单元格内容解析为浮点数
					if value, err := strconv.ParseFloat(cell, 64); err == nil {
						// 只有在值是小数的情况下才改变格式
						if value != float64(int(value)) { // 判断是否为小数
							entry[strings.ToLower(headers[i])] = fmt.Sprintf("%.4f", value) // 格式化为两位小数
						} else {
							entry[strings.ToLower(headers[i])] = cell // 保持整数原样
						}
					} else {
						// 不是数值，直接使用字符串
						entry[strings.ToLower(headers[i])] = cell
					}
				}
				if ikey != "-1" || len(entry) > 0 {
					data[ikey] = entry
				}
			}
		}

		// 将数据写入 JSON 文件
		jsonData, err := json.MarshalIndent(data, "", "  ")
		if err != nil {
			log.Printf("无法生成 JSON 数据: %v", err)
			continue
		}

		err = os.WriteFile(jsonFile, jsonData, 0644)
		if err != nil {
			log.Printf("无法写入 JSON 文件 %s: %v", jsonFile, err)
			continue
		}

		fmt.Printf("转换成功，输出文件: %s\n", jsonFile)
	}
}
