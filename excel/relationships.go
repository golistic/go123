// Copyright (c) 2023, Geert JM Vanderkelen

package excel

import (
	"encoding/xml"
)

type Relationships struct {
	XMLName       xml.Name        `xml:"Relationships"`
	Relationships []*Relationship `xml:"Relationship"`
}

func (rels *Relationships) GetID(id string) *Relationship {
	for _, rel := range rels.Relationships {
		if rel.ID == id {
			return rel
		}
	}

	return nil
}

type Relationship struct {
	XMLName xml.Name `xml:"Relationship"`
	ID      string   `xml:"Id,attr"`
	Type    string   `xml:"Type,attr"`
	Target  string   `xml:"Target,attr"`
}
