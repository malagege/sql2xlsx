package sql2xlsx

// A simple example of quickly converting SQL result into an Excel file.
//參考：https://github.com/bwmarrin/sql2xlsx
import (
	"database/sql"
	"fmt"
	"time"

	"github.com/tealeg/xlsx"
)

// Variables used with command line arguments
var (
	host string
	user string
	pass string
	sqlf string
	outf string
)

func GenerateXLSXFromRows(rows *sql.Rows, outf string, writeheader bool) error {

	var err error

	// Get column names from query result
	colNames, err := rows.Columns()
	if err != nil {
		return fmt.Errorf("error fetching column names, %s\n", err)
	}
	length := len(colNames)

	// Create a interface slice filled with pointers to interface{}'s
	pointers := make([]interface{}, length)
	container := make([]interface{}, length)
	for i := range pointers {
		pointers[i] = &container[i]
	}

	// Create output xlsx workbook
	xfile := xlsx.NewFile()
	xsheet, err := xfile.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("error adding sheet to xlsx file, %s\n", err)
	}

	// Write Headers to 1st row
	var xrow *xlsx.Row
	if writeheader == true {
		xrow = xsheet.AddRow()
		xrow.WriteSlice(&colNames, -1)
	}
	// Process sql rows
	for rows.Next() {

		// Scan the sql rows into the interface{} slice
		err = rows.Scan(pointers...)
		if err != nil {
			return fmt.Errorf("error scanning sql row, %s\n", err)
		}

		xrow = xsheet.AddRow()

		// Here we range over our container and look at each column
		// and set some different options depending on the column type.
		for _, v := range container {
			xcell := xrow.AddCell()
			switch v := v.(type) {
			case string:
				xcell.SetString(v)
			case []byte:
				xcell.SetString(string(v))
			case int64:
				xcell.SetInt64(v)
			case float64:
				xcell.SetFloat(v)
			case bool:
				xcell.SetBool(v)
			case time.Time:
				xcell.SetDateTime(v)
			default:
				xcell.SetValue(v)
			}

		}

	}

	// Save the excel file to the provided output file
	err = xfile.Save(outf)
	if err != nil {
		return fmt.Errorf("error writing to output file %s, %s\n", outf, err)
	}

	return nil
}
