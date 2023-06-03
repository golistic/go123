// Copyright (c) 2023, Geert JM Vanderkelen

package excel_test

import (
	"errors"
	"fmt"
	"testing"

	"github.com/golistic/xt"

	"github.com/golistic/go123/excel"
)

func TestNew(t *testing.T) {
	t.Run("cannot open other kinds of Open XML files", func(t *testing.T) {
		_, err := excel.New("_testdata/word.docx")
		xt.KO(t, err)
		xt.Assert(t, errors.Is(err, excel.ErrNotMSExcel))
	})

	t.Run("cannot open other kinds of Open XML files", func(t *testing.T) {
		_, err := excel.New("_testdata/testsheets.xlsx")
		xt.OK(t, err)
	})
}

func TestOpenXML_Sheets(t *testing.T) {
	ox, err := excel.New("_testdata/testsheets.xlsx")
	xt.OK(t, err)
	defer ox.MustClose()

	t.Run("getting sheet information", func(t *testing.T) {
		sheets, err := ox.Sheets()
		xt.OK(t, err)
		xt.Eq(t, 3, len(sheets))
		xt.Eq(t, "dogs", sheets[1].Name)
		xt.Eq(t, "worksheets/sheet2.xml", sheets[1].Target)
	})
}

func TestOpenXML_Worksheet(t *testing.T) {
	ox, err := excel.New("_testdata/testsheets.xlsx")
	xt.OK(t, err)
	defer ox.MustClose()

	t.Run("get dogs worksheet", func(t *testing.T) {
		ws, err := ox.Worksheet("dogs")
		xt.OK(t, err)

		for _, row := range ws.SheetData.Rows {
			for _, cell := range row.Cells {
				fmt.Println("Cell value:", cell.Value)
			}
		}
	})
}
