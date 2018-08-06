package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
	"time"
)

var fileName string
var desc string
var date string
var income bool
var amount float64
var year int
var month string
var day int

func init() {

	// Get the current date
	currentYear, cmonth, currentDay := time.Now().Date()

	// Convert the current month to a String
	currentMonth := cmonth.String()

	// CONSOLE FLAGS //
	flag.BoolVar(&income, "i", false, "If flag is present, then it is income. If not, then it is expense.")
	flag.StringVar(&desc, "de", "No information provided", "Description of the expense")
	flag.IntVar(&currentDay, "d", currentDay, "Day of the income or expense")
	flag.StringVar(&currentMonth, "m", currentMonth, "Month of the income or expense")
	flag.IntVar(&currentYear, "y", currentYear, "Year of the income or expense")
	flag.Float64Var(&amount, "a", 0.00, "Amount of the expense in $ to the nearest hundreth.")
	flag.Parse()

	// If the flags are used to change the dates, then the new dates will be used
	// If not, use the current date
	day = currentDay
	month = currentMonth
	year = currentYear

	// File name format: FinanceYYYY.xlsx
	fileName = "Finance" + strconv.Itoa(year) + ".xlsx"

}

func main() {
	date := strconv.Itoa(day) + "-" + month + "-" + strconv.Itoa(year)
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		xlsx = excelize.NewFile()
		xlsx.SetSheetName("Sheet1", "January")
		xlsx.NewSheet("February")
		xlsx.NewSheet("March")
		xlsx.NewSheet("April")
		xlsx.NewSheet("May")
		xlsx.NewSheet("June")
		xlsx.NewSheet("July")
		xlsx.NewSheet("August")
		xlsx.NewSheet("September")
		xlsx.NewSheet("October")
		xlsx.NewSheet("November")
		xlsx.NewSheet("December")

		titleStyle, _ := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#7c7a87"],"pattern":5}}`)

		for i := 1; i <= xlsx.SheetCount; i++ {
			xlsx.SetCellStyle(xlsx.GetSheetName(i), "A1", "C1", titleStyle)
			xlsx.SetCellValue(xlsx.GetSheetName(i), "A1", "Expense Description")
			xlsx.SetCellValue(xlsx.GetSheetName(i), "B1", "Amount")
			xlsx.SetCellValue(xlsx.GetSheetName(i), "C1", "Date")
			xlsx.SetCellValue(xlsx.GetSheetName(i), "E1", "Total")
			xlsx.SetColWidth(xlsx.GetSheetName(i), "A", "A", 30)
			xlsx.SetColWidth(xlsx.GetSheetName(i), "B", "C", 20)

			xlsx.SetCellFormula(xlsx.GetSheetName(i), "E3", "=SUM(B3:B999)")

		}

	}

	if !income {
		amount *= -1
	}

	xlsx.SetActiveSheet(xlsx.GetSheetIndex(month))

	rows := xlsx.GetRows(month)

	nextRow := len(rows) + 1

	fmt.Println(nextRow)

	data := []interface{}{desc, amount, date}

	xlsx.SetSheetRow(month, "A"+strconv.Itoa(nextRow), &data)

	amountStyle, _ := xlsx.NewStyle(`{"number_format": 177}`)
	xlsx.SetCellStyle(month, "B"+strconv.Itoa(nextRow), "B"+strconv.Itoa(nextRow), amountStyle)
	xlsx.SetCellStyle(month, "E3", "E3", amountStyle)

	dateStyle, _ := xlsx.NewStyle(`{"number_format": 22}`)
	xlsx.SetCellStyle(month, "C"+strconv.Itoa(nextRow), "C"+strconv.Itoa(nextRow), dateStyle)

	xlsx.SaveAs(fileName)
}
