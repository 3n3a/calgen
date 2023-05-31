package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/arran4/golang-ical"
	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
)

const (
	DEFAULT_PROD_ID = "-//3n3a//calgen//EN"
	DEFAULT_CAL_EXTENSION = "ics"
)

type DateOrEmpty struct {
	TS            time.Time
	Empty 		  bool
}

type Meeting struct {
	Empty         bool
	Date          DateOrEmpty
	Name          string
	Place         string
	Subject       string
	Goals         string
	MeetingPlace  string
	TimeTableInfo string
	MainPerson    string
	OtherPerson1  string
	OtherPerson2  string
}

func daysSinceEpochToTime(ts int) time.Time {
	// start time, according to excel
	d := time.Date(1900, 1, 1, 0, 0, 0, 0, time.Local)

	// two because bug in excel...
	newDate := d.AddDate(0, 0, ts-2)
	return newDate
}

func getAtPosOrDefault[T any](arr []T, at int64, def T) T {
	if int64(len(arr)-1) < at {
		return def
	}
	return arr[at]
}

func parseMeeting(rowValues []string) (Meeting, bool) {
	isEmpty := len(rowValues) < 1
	m := new(Meeting)

	rawDateValue := getAtPosOrDefault(rowValues, 0, "")
	daysSince1900, err := strconv.Atoi(rawDateValue)
	m.Date = DateOrEmpty{
		TS: daysSinceEpochToTime(daysSince1900),
		Empty: rawDateValue == "" || err != nil,
	}
	m.Name = getAtPosOrDefault(rowValues, 1, "")
	m.Place = getAtPosOrDefault(rowValues, 2, "")
	m.Subject = getAtPosOrDefault(rowValues, 3, "")
	m.Goals = getAtPosOrDefault(rowValues, 4, "")
	m.MeetingPlace = getAtPosOrDefault(rowValues, 5, "")
	m.TimeTableInfo = getAtPosOrDefault(rowValues, 6, "")
	m.MainPerson = getAtPosOrDefault(rowValues, 7, "")
	m.OtherPerson1 = getAtPosOrDefault(rowValues, 8, "")
	m.OtherPerson2 = getAtPosOrDefault(rowValues, 9, "")
	return *m, isEmpty
}

// Parse an Excel Sheet and output as an array of your Type
//
// You need to proovide a `parseElement` function that accepts an array of strings (the values from the row).
// This function should then return a filled out struct of type T
func parseExcelSheet[T any](file string, sheet string, parseElement func(rowValues []string) (T, bool)) []T {
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

		// fmt.Printf("%+v\n", row)
		// Parse the values and assign to the struct fields
		element, isEmpty := parseElement(row)

		if !isEmpty {
			elements = append(elements, element)
		}
	}

	return elements
}

func createMeetingCalendar(meetings []Meeting, eventTitle string, inlcudeEmptyOnes bool) (string) {
	cal := ics.NewCalendar()
	cal.SetProductId(DEFAULT_PROD_ID)
	cal.SetMethod(ics.Method(ics.ComponentVCalendar))

	for _, m := range meetings {
		if ((m.Empty == false && m.Date.Empty == false) || inlcudeEmptyOnes) {
			ev := cal.AddEvent(fmt.Sprintf("calgen-%s", uuid.New().String()))
			ev.SetDtStampTime(m.Date.TS)
			ev.SetAllDayStartAt(m.Date.TS)
			ev.SetSummary(fmt.Sprintf("%s - %s - %s", eventTitle, m.Name, m.Subject))
			ev.SetDescription(fmt.Sprintf("Ziele:\n%s\n\nTreffpunkt:\n%s\n\nDauer / Rückkehr:\n%s\n\nLeiter: %s, %s, %s\n", m.Goals, m.MeetingPlace, m.TimeTableInfo, m.MainPerson, m.OtherPerson1, m.OtherPerson2))
			ev.SetLocation(m.Place)
		}
	}
	return cal.Serialize()
}

func saveAsCalendarFile(filename string, calendarString string) error {
	f, err := os.Create(fmt.Sprintf("%s.%s", filename, DEFAULT_CAL_EXTENSION))
	if err != nil {
		return err
	}
	defer f.Close()

	_, err = f.WriteString(calendarString)
	if err != nil {
		return err
	}

	return nil
}

func getMeetingsForPerson(person string, filename string, sheet string, eventTitlePrefix string) []Meeting {
	var personMeetings []Meeting
	meetings := parseExcelSheet[Meeting](filename, sheet, parseMeeting)

	// Search for Person in MainPerson, OtherPerson1, OtherPerson2
	for _, meeting := range meetings {
		if (strings.Contains(meeting.MainPerson, person) || strings.Contains(meeting.OtherPerson1, person) || strings.Contains(meeting.OtherPerson2, person)) {
			personMeetings = append(personMeetings, meeting)
		}
	}

	// // Debug Output of all Meetings
	// oj, _ := json.MarshalIndent(selected, "", "  ")
	// fmt.Printf("%s\n", string(oj))
	return personMeetings
}



func generateCalendarForPerson() {
	

	inputFile := "ol.xlsx"
	sheetName := "alle Anlässe"
	eventTitlePrefix := "OL Training"
	person := "Enea"
	

	meetings := getMeetingsForPerson(person, inputFile, sheetName, eventTitlePrefix)
	personCalendar := createMeetingCalendar(meetings, eventTitlePrefix, false)
	err := saveAsCalendarFile(strings.ToLower(person), personCalendar)
	if err != nil {
		panic(err)
	}
}


func generateCalendarForAllEvents() {
	
	inputFile := "ol.xlsx"
	sheetName := "alle Anlässe"
	eventTitlePrefix := "OL Training"
	outFileName := "trainings"
	

	meetings := parseExcelSheet[Meeting](inputFile, sheetName, parseMeeting)
	personCalendar := createMeetingCalendar(meetings, eventTitlePrefix, false)
	err := saveAsCalendarFile(strings.ToLower(outFileName), personCalendar)
	if err != nil {
		panic(err)
	}
}

func main()  {
	generateCalendarForAllEvents()
}