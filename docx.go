package office

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"github.com/beevik/etree"
	"io"
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

func DocxToString(docx []byte) (string, error) {
	newBuffer := bytes.NewReader(docx)
	reader, err := zip.NewReader(newBuffer, int64(len(docx)))
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
		if e.Tag == "br" {
			sum.WriteString("\n")
		} else if e.Tag == "p" {
			sum.WriteString("\n")
		} else {
			if text := e.Text(); text != "" {
				sum.WriteString(text)
			}
			if value := getText(e); value != "" {
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
			//text = strings.ReplaceAll(text, " ", "")
			//text = strings.ReplaceAll(text, "\n", "")
			//text = strings.ReplaceAll(text, "\t", "")
			//if text != "" {
			//	arr = append(arr, text)
			//}
			arr = append(arr, text)
		}
		if value := getTextArray(e); len(value) != 0 {
			arr = append(arr, value...)
		}
	}
	return arr
}

func WordSplitter(text string, maxLength int) []string {
	var chunks []string
	words := strings.FieldsFunc(text, func(r rune) bool {
		return r == '!' || r == '?' || r == ';' || r == '。'
	})
	var currentChunk []string
	currentLength := 0

	/*for _, word := range words {
		if currentLength+len(word)+1 > maxLength {
			chunks = append(chunks, strings.Join(currentChunk, " "))
			currentChunk = []string{word}
			currentLength = len(word)
		} else {
			currentChunk = append(currentChunk, word)
			currentLength += len(word) + 1
		}
	}*/

	for _, word := range words {
		if currentLength <= maxLength-1 {
			currentChunk = append(currentChunk, word)
			currentLength += 1
		} else {
			currentLength = 0
			chunks = append(chunks, strings.Join(currentChunk, ""))
			currentChunk = nil
		}
	}
	if len(currentChunk) > 0 {
		chunks = append(chunks, strings.Join(currentChunk, ""))
	}
	return chunks
}
