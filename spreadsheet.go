package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

type Employee struct {
	ID       int
	Name     string
	Position string
}

// Parse an Excel Sheet and output as an array of your Type
//
// You need to proovide a `parseElement` function that accepts an array of strings (the values from the row).
// This function should then return a filled out struct of type T
func parseExcelSheet[T interface](file string, sheet string, parseElement func(rowValues []string) T) []T {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println("Failed to open Excel file:", err)
		return
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("Failed to get rows from the sheet:", err)
		return
	}

	var elements []T
	for i, row := range rows {
		// Skip the header row
		if i == 0 {
			continue
		}

		// Parse the values and assign to the struct fields
		element := parseElement(row)

		elements = append(elements, element)
	}

    return elements
}
