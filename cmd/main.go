package main

import (
	"fmt"
	"github.com/jimu-server/office"
	"os"
)

func main() {
	var err error
	var text string
	var file *os.File
	if file, err = os.Open("D:\\Code\\go_code\\jimu-server\\office\\cmd\\test.docx"); err != nil {
		fmt.Println(err.Error())
		return
	}
	text, err = office.DocxToString(file)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	fmt.Println(text)
	fmt.Println(len(text))
}
