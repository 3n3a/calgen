package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	_ "strconv"
)

type Meeting struct {
	Date       string
	Name     string
	Place string
	Subject string
	Goals string
	MeetingPlace string
	TimeTableInfo string
	MainPerson string
	OtherPerson1 string
	OtherPerson2 string
}

func parseMeeting(rowValues []string) Meeting {
	m := new(Meeting)
	m.Date = rowValues[0]
	m.Name = rowValues[1]
	m.Place = rowValues[2]
	m.Subject = rowValues[3]
	m.Goals = rowValues[4]
	m.MeetingPlace = rowValues[5]
	m.TimeTableInfo = rowValues[6]
	m.MainPerson = rowValues[7]
	m.OtherPerson1 = rowValues[8]
	m.OtherPerson2 = rowValues[9]
	return *m
}

// Parse an Excel Sheet and output as an array of your Type
//
// You need to proovide a `parseElement` function that accepts an array of strings (the values from the row).
// This function should then return a filled out struct of type T
func parseExcelSheet[T any](file string, sheet string, parseElement func(rowValues []string) T) []T {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println("Failed to open Excel file:", err)
		return nil
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("Failed to get rows from the sheet:", err)
		return nil
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

func main() {
    o := parseExcelSheet[Meeting]("ol.xlsx", "alle Anlässe", parseMeeting)
	fmt.Printf("%+v\n", o)
}