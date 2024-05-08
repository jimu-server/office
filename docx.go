package office

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"github.com/beevik/etree"
	"io"
	"os"
	"strings"
)

type XMLNode struct {
	XMLName xml.Name
	Content string `xml:",innerxml"`
}

type XMLBody struct {
	Content []XMLNode `xml:",any"`
}

type Docx struct {
	XMLName xml.Name `xml:"document"`
	Body    XMLBody  `xml:"body"`
}

func DocxToString(docx *os.File) (string, error) {
	stat, err := docx.Stat()
	if err != nil {
		return "", err
	}
	reader, err := zip.NewReader(docx, stat.Size())
	if err != nil {
		return "", err
	}
	var rc io.ReadCloser
	var buffer bytes.Buffer
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			rc, err = file.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()
			_, err = io.Copy(&buffer, rc)
			if err != nil {
				return "", err
			}
			break
		}
	}

	document := etree.NewDocument()
	if err = document.ReadFromBytes(buffer.Bytes()); err != nil {
		return "", err
	}
	root := document.Root()

	return getText(root), nil
}

func getText(element *etree.Element) string {
	sum := strings.Builder{}
	if element == nil {
		return ""
	}
	for _, e := range element.ChildElements() {
		if text := e.Text(); text != "" {
			text = strings.ReplaceAll(text, " ", "")
			text = strings.ReplaceAll(text, "\n", "")
			text = strings.ReplaceAll(text, "\t", "")
			if text != "" {
				sum.WriteString(text)
			}
		}
		if value := getText(e); value != "" {
			value = strings.ReplaceAll(value, " ", "")
			value = strings.ReplaceAll(value, "\n", "")
			value = strings.ReplaceAll(value, "\t", "")
			if value != "" {
				sum.WriteString(value)
			}
		}
	}
	return sum.String()
}

func DocxToStringSlice(docx []byte) ([]string, error) {
	newBuffer := bytes.NewReader(docx)
	reader, err := zip.NewReader(newBuffer, int64(len(docx)))
	if err != nil {
		return nil, err
	}
	var rc io.ReadCloser
	var buffer bytes.Buffer
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			rc, err = file.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			_, err = io.Copy(&buffer, rc)
			if err != nil {
				return nil, err
			}
			break
		}
	}

	document := etree.NewDocument()
	if err = document.ReadFromBytes(buffer.Bytes()); err != nil {
		return nil, err
	}
	root := document.Root()

	return getTextArray(root), nil
}

func getTextArray(element *etree.Element) []string {
	arr := make([]string, 0)
	if element == nil {
		return arr
	}
	for _, e := range element.ChildElements() {
		if text := e.Text(); text != "" {
			text = strings.ReplaceAll(text, " ", "")
			text = strings.ReplaceAll(text, "\n", "")
			text = strings.ReplaceAll(text, "\t", "")
			if text != "" {
				arr = append(arr, text)
			}
		}
		if value := getTextArray(e); len(value) != 0 {
			arr = append(arr, value...)
		}
	}
	return arr
}
