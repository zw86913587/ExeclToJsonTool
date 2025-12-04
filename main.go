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
	excelFiles, err := findExcelFiles()
	if err != nil {
		log.Fatalf("查找Excel文件时出错: %v", err)
	}

	if len(excelFiles) == 0 {
		fmt.Println("当前目录中未找到任何 Excel 文件 (.xlsx 或 .xls)")
		return
	}

	fmt.Printf("找到 %d 个Excel文件:\n", len(excelFiles))
	for _, file := range excelFiles {
		fmt.Printf("  - %s\n", file)
	}

	// 使用 WaitGroup 等待所有 goroutine 完成
	var wg sync.WaitGroup
	// 创建一个通道用于接收转换结果
	resultChan := make(chan string, len(excelFiles))
	errorsChan := make(chan error, len(excelFiles))

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
			if err := convertExcelToJSON(file); err != nil {
				errorsChan <- fmt.Errorf("转换文件 %s 时出错: %v", file, err)
			} else {
				resultChan <- fmt.Sprintf("转换成功，文件: %s", file)
			}
		}(excelFile)
	}

	// 等待所有 goroutine 完成
	wg.Wait()
	close(resultChan)
	close(errorsChan)

	// 输出所有错误信息
	fmt.Println("\n=== 转换结果 ===")
	if errCount := len(errorsChan); errCount > 0 {
		fmt.Printf("有 %d 个文件转换失败:\n", errCount)
		for err := range errorsChan {
			log.Printf("错误: %v", err)
		}
	} else {
		fmt.Println("所有文件都转换成功!")
	}

	fmt.Println("所有文件处理完成.")
}

func findExcelFiles() ([]string, error) {
	currentDir, err := os.Getwd()
	if err != nil {
		return nil, fmt.Errorf("无法获取当前目录: %v", err)
	}

	var excelFiles []string
	err = filepath.Walk(currentDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() &&
			(filepath.Ext(info.Name()) == ".xlsx" || filepath.Ext(info.Name()) == ".xls") &&
			!strings.Contains(info.Name(), "~") {
			excelFiles = append(excelFiles, path)
		}
		return nil
	})

	return excelFiles, err
}

func convertExcelToJSON(excelFile string) error {
	fmt.Printf("开始处理文件: %s\n", excelFile)

	jsonDir := filepath.Join(filepath.Dir(excelFile), "json")

	if err := os.MkdirAll(jsonDir, os.ModePerm); err != nil {
		return fmt.Errorf("无法创建文件夹 %s: %v", jsonDir, err)
	}

	fileName := filepath.Base(excelFile)
	jsonFile := filepath.Join(jsonDir, strings.TrimSuffix(fileName, filepath.Ext(fileName))+".json")

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("无法打开 Excel 文件 %s: %v", excelFile, err)
	}
	defer func() {
		if closeErr := f.Close(); closeErr != nil {
			fmt.Printf("警告: 关闭文件 %s 时出错: %v\n", excelFile, closeErr)
		}
	}()

	sheetName := f.GetSheetName(0)
	fmt.Printf("正在读取工作表: %s\n", sheetName)

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("无法读取工作表 %s: %v", sheetName, err)
	}

	// 确保有足够的行数进行处理
	if len(rows) < 4 {
		return fmt.Errorf("Excel文件行数不足，至少需要4行数据，当前只有 %d 行", len(rows))
	}

	fmt.Printf("Excel文件共有 %d 行数据\n", len(rows))

	// 检查第四行标记
	markRow := rows[3]
	fmt.Printf("第四行标记内容有 %d 列\n", len(markRow))

	data := make(map[string]map[string]string)
	headers := rows[0]

	// 显示表头
	fmt.Printf("表头有 %d 列: %v\n", len(headers), headers)

	for rowIndex, row := range rows[4:] {
		entry := make(map[string]string)
		key := ""

		// 补全空白格子
		if len(markRow) > len(row) {
			padding := make([]string, len(markRow)-len(row))
			row = append(row, padding...)
		}

		processedCols := 0
		for colIndex, cell := range row {
			// 确保不越界 - 这是关键修复点
			if colIndex >= len(markRow) {
				// 当前行的数据列超出了标记行的长度，跳过额外的列
				break
			}

			// 只处理标记为"3"的列
			if markRow[colIndex] != "3" {
				continue
			}

			processedCols++

			// 确保表头存在
			if colIndex >= len(headers) {
				continue
			}

			// 第一列作为键值
			if colIndex == 0 {
				key = cell
			}

			// 处理数值，保留适当的小数位
			if value, err := strconv.ParseFloat(cell, 64); err == nil {
				if value != float64(int(value)) {
					entry[strings.ToLower(headers[colIndex])] = fmt.Sprintf("%.4f", value)
				} else {
					entry[strings.ToLower(headers[colIndex])] = cell
				}
			} else {
				entry[strings.ToLower(headers[colIndex])] = cell
			}
		}

		// 只有当有数据或者设置了key时才添加到结果中
		if key != "" || len(entry) > 0 {
			if key == "" {
				key = fmt.Sprintf("row_%d", rowIndex)
			}
			data[key] = entry
			fmt.Printf("处理第 %d 行数据，键值: %s，处理了 %d 列\n", rowIndex+5, key, processedCols)
		}
	}

	fmt.Printf("总共处理了 %d 条数据记录\n", len(data))

	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return fmt.Errorf("无法生成 JSON 数据: %v", err)
	}

	if err = os.WriteFile(jsonFile, jsonData, 0644); err != nil {
		return fmt.Errorf("无法写入 JSON 文件 %s: %v", jsonFile, err)
	}

	fmt.Printf("成功写入JSON文件: %s\n", jsonFile)

	return nil
}
