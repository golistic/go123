// Copyright (c) 2023, Geert JM Vanderkelen

package excel

import "encoding/xml"

type Worksheet struct {
	XMLName   xml.Name  `xml:"worksheet"`
	SheetData SheetData `xml:"sheetData"`
}

type SheetData struct {
	XMLName xml.Name `xml:"sheetData"`
	Rows    []Row    `xml:"row"`
}

type Row struct {
	XMLName xml.Name `xml:"row"`
	Cells   []Cell   `xml:"c"`
}

type Cell struct {
	XMLName xml.Name `xml:"c"`
	Value   string   `xml:"v"`
}
