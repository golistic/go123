// Copyright (c) 2023, Geert JM Vanderkelen

package excel

import (
	"encoding/xml"
	"fmt"
	"io"
)

type Workbook struct {
	XMLName xml.Name `xml:"workbook"`
	Sheets  []*Sheet `xml:"sheets>sheet"`
}

func newWorkbook(rc io.ReadCloser) (*Workbook, error) {
	var wb Workbook
	err := xml.NewDecoder(rc).Decode(&wb)
	if err != nil {
		return nil, fmt.Errorf("excel: parsing the workbook XML (%w)", err)
	}

	return &wb, nil
}

type Sheet struct {
	XMLName    xml.Name `xml:"sheet"`
	Name       string   `xml:"name,attr"`
	SheetID    int      `xml:"sheetId,attr"`
	RelationID string   `xml:"id,attr"`

	Target string
}
