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
	"sync"
)

func main() {
	excelFiles, done := findExcelFiles()
	if done {
		return
	}

	// 使用 WaitGroup 等待所有 goroutine 完成
	var wg sync.WaitGroup
	// 创建一个通道用于接收转换结果
	resultChan := make(chan string)

	// 启动一个 goroutine 来处理结果
	go func() {
		for result := range resultChan {
			fmt.Println(result)
		}
	}()

	// 遍历找到的每个 Excel 文件并启动 goroutine
	for _, excelFile := range excelFiles {
		wg.Add(1)
		go func(file string) {
			defer wg.Done()
			convertExcelToJSON(file, resultChan)
		}(excelFile)
	}

	// 等待所有 goroutine 完成
	wg.Wait()
	close(resultChan)
	fmt.Println("按回车键继续...")
	_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
}

func findExcelFiles() ([]string, bool) {
	currentDir, err := os.Getwd()
	if err != nil {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		log.Fatalf("无法获取当前目录: %v", err)
	}

	var excelFiles []string
	err = filepath.Walk(currentDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			fmt.Println("按回车键继续...")
			_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
			return err
		}
		if !info.IsDir() && (filepath.Ext(info.Name()) == ".xlsx" || filepath.Ext(info.Name()) == ".xls") && strings.Index(info.Name(), "~") == -1 {
			excelFiles = append(excelFiles, path)
		}
		return nil
	})

	if err != nil {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		log.Fatalf("遍历目录时出错: %v", err)
	}

	if len(excelFiles) == 0 {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		fmt.Println("当前目录中未找到任何 Excel 文件 (.xlsx 或 .xls)")
		return nil, true
	}
	return excelFiles, false
}

func convertExcelToJSON(excelFile string, resultChan chan string) {
	jsonDir := filepath.Join(filepath.Dir(excelFile), "json")

	if err := os.MkdirAll(jsonDir, os.ModePerm); err != nil {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		log.Printf("无法创建文件夹 %s: %v", jsonDir, err)
		return
	}

	fileName := filepath.Base(excelFile)
	jsonFile := filepath.Join(jsonDir, fileName[:len(fileName)-len(filepath.Ext(fileName))]+".json")

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		log.Printf("无法打开 Excel 文件 %s: %v", excelFile, err)
		return
	}

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		log.Printf("无法读取工作表 %s: %v", sheetName, err)
		return
	}

	data := make(map[string]map[string]string)
	if len(rows) > 0 {
		headers := rows[0]
		for _, row := range rows[4:] {
			entry := make(map[string]string)
			ikey := "-1"
			//判断有多少个列需要转换
			m := len(rows[3])
			//for k := range rows[3] {
			//	if rows[3][k] == "3" {
			//		m++
			//	}
			//}
			//补全空白格子
			if m > len(row) {
				j := 0
				n := m - len(row)
				for j <= n {
					row = append(row, "")
					j++
				}
			}
			for i, cell := range row {
				if i+1 > len(rows[3]) {
					break
				}
				if rows[3][i] != "3" {
					continue
				}
				if i == 0 {
					ikey = cell
				}
				if value, err := strconv.ParseFloat(cell, 64); err == nil {
					if value != float64(int(value)) {
						entry[strings.ToLower(headers[i])] = fmt.Sprintf("%.4f", value)
					} else {
						entry[strings.ToLower(headers[i])] = cell
					}
				} else {
					entry[strings.ToLower(headers[i])] = cell
				}
			}
			if ikey != "-1" || len(entry) > 0 {
				data[ikey] = entry
			}
		}
	}

	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		log.Printf("无法生成 JSON 数据: %v", err)
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		return
	}

	err = os.WriteFile(jsonFile, jsonData, 0644)
	if err != nil {
		log.Printf("无法写入 JSON 文件 %s: %v", jsonFile, err)
		fmt.Println("按回车键继续...")
		_, _ = fmt.Scanln() // 暂停程序，等待用户按回车键
		return
	}

	resultChan <- fmt.Sprintf("转换成功，输出文件: %s", jsonFile)
}
