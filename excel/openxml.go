// Copyright (c) 2023, Geert JM Vanderkelen

package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path"
)

// OpenXML defines an Excel Open XML file.
type OpenXML struct {
	filename    string
	rcZipCloser *zip.ReadCloser
	rZip        *zip.Reader
	origReader  *io.Reader
}

// New takes filename as the location of an Excel Open XML file and tries
// to open it. This will fail for any other Open XML files and return error
// ErrNotMSExcel.
// The caller must call OpenXML.Close() when reader is no longer needed.
func New(filename string) (*OpenXML, error) {
	ox := &OpenXML{
		filename: filename,
	}

	if err := ox.openZipFile(); err != nil {
		return nil, err
	}

	if err := ox.verifyExcel(); err != nil {
		return nil, err
	}

	return ox, nil
}

// NewWithReader opens the Excel Open XML file for reading using an
// io.Reader. The size argument is the size of the ZIP-file.
func NewWithReader(r io.Reader, size int64) (*OpenXML, error) {
	ox := &OpenXML{}

	if err := ox.openZipFileWithReader(r, size); err != nil {
		return nil, err
	}

	if err := ox.verifyExcel(); err != nil {
		return nil, err
	}

	return ox, nil
}

// Close the reader. This must be called when reader is no longer needed.
// When ox was created using an io.Reader, Close does not do anything.
func (ox *OpenXML) Close() error {
	if ox.rcZipCloser != nil {
		return ox.rcZipCloser.Close()
	}

	return nil
}

// MustClose the reader but panics on error instead. This is mainly useful for
// tests. See Close.
func (ox *OpenXML) MustClose() {
	if ox.rcZipCloser != nil {
		if err := ox.rcZipCloser.Close(); err != nil {
			panic(fmt.Sprintf("excel: %s", err))
		}
	}
}

func (ox *OpenXML) Sheets() ([]*Sheet, error) {
	rc, err := ox.openWorkbookFile()
	if err != nil {
		return nil, err
	}
	defer func() { _ = rc.Close() }()

	wb, err := newWorkbook(rc)
	if err != nil {
		return nil, err
	}

	rels, err := ox.getWorkbookRelationships()
	if err != nil {
		return nil, err
	}

	targets := map[string]*Relationship{}
	for _, rel := range rels.Relationships {
		targets[rel.ID] = rel
	}

	for _, s := range wb.Sheets {
		s.Target = targets[s.RelationID].Target
	}

	return wb.Sheets, nil
}

func (ox *OpenXML) Worksheet(name string) (*Worksheet, error) {
	rc, err := ox.openWorksheetFile(name)
	if err != nil {
		return nil, err
	}
	defer func() { _ = rc.Close() }()

	var worksheet Worksheet
	err = xml.NewDecoder(rc).Decode(&worksheet)
	if err != nil {
		return nil, fmt.Errorf("excel: parsing worksheet '%s' (%w)", name, err)
	}

	return &worksheet, err
}

func (ox *OpenXML) verifyExcel() error {

	r, err := ox.getRelationships("_rels/.rels")
	if err != nil || r.GetID("rId1").Target != "xl/workbook.xml" {
		return fmt.Errorf("excel: opening (%w)", ErrNotMSExcel)
	}

	return nil
}

// openZipFile opens the Excel Open XML file for reading.
func (ox *OpenXML) openZipFile() error {

	var err error
	ox.rcZipCloser, err = zip.OpenReader(ox.filename)
	if err != nil {
		return fmt.Errorf("excel: opening file '%s' (%w)", ox.filename, err)
	}

	return nil
}

// openZipFileWithReader opens the Excel Open XML file for reading using an
// io.Reader. The size argument is the size of the ZIP-file.
func (ox *OpenXML) openZipFileWithReader(r io.Reader, size int64) error {

	rZip, err := zip.NewReader(r.(io.ReaderAt), size)
	if err != nil {
		return fmt.Errorf("excel: opening zip file (%w)", err)
	}

	ox.rZip = rZip

	return nil
}

func (ox *OpenXML) openFile(filename string) (io.ReadCloser, error) {
	if ox.rZip != nil {
		for _, f := range ox.rZip.File {
			if f.Name == filename {
				return f.Open()
			}
		}
	}

	if ox.rcZipCloser != nil {
		for _, f := range ox.rcZipCloser.File {
			if f.Name == filename {
				return f.Open()
			}
		}
	}

	return nil, fmt.Errorf("excel: opening file '%s' (not available)", filename)
}

func (ox *OpenXML) getFileFromZip(filePath string) (*zip.File, error) {

	if ox.rcZipCloser != nil {
		for _, f := range ox.rcZipCloser.File {
			if f.Name == filePath {
				return f, nil
			}
		}
	}

	if ox.rZip != nil {
		for _, f := range ox.rZip.File {
			if f.Name == filePath {
				return f, nil
			}
		}
	}

	return nil, fmt.Errorf("getting file %s", filePath)
}

// openWorksheet looks up worksheet with given name and returns the
// opened file.
// The caller must close the reader when no longer needed.
func (ox *OpenXML) openWorksheetFile(name string) (io.ReadCloser, error) {
	const root = "xl"

	sheets, err := ox.Sheets()
	if err != nil {
		return nil, err
	}

	var sheet *Sheet
	for _, s := range sheets {
		if s.Name == name {
			sheet = s
			break
		}
	}

	if sheet == nil {
		return nil, fmt.Errorf("excel: getting sheet information (%w)", ErrSheetNotAvailable)
	}

	rels, err := ox.getWorkbookRelationships()
	if err != nil {
		return nil, err
	}

	var target string
	if rel := rels.GetID(sheet.RelationID); rel != nil {
		target = path.Join(root, rel.Target)
	}

	if target == "" {
		return nil, fmt.Errorf("excel: getting sheet information (relationship not available)")
	}

	return ox.openFile(target)
}

func (ox *OpenXML) openWorkbookFile() (io.ReadCloser, error) {
	const wbPath = "xl/workbook.xml"

	rc, err := ox.openFile(wbPath)
	if err != nil {
		return nil, fmt.Errorf("excel: opening workbook (%w)", err)
	}

	return rc, nil
}

func (ox *OpenXML) getRelationships(filename string) (*Relationships, error) {
	rc, err := ox.openFile(filename)
	if err != nil {
		return nil, err
	}
	defer func() { _ = rc.Close() }()

	var rels Relationships
	if err := xml.NewDecoder(rc).Decode(&rels); err != nil {
		return nil, fmt.Errorf("excel: parsing the workbook relationship XML (%w)", err)
	}

	return &rels, nil
}

func (ox *OpenXML) getWorkbookRelationships() (*Relationships, error) {
	return ox.getRelationships("xl/_rels/workbook.xml.rels")
}
