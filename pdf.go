package office

import (
	"bytes"
	"github.com/pdfcpu/pdfcpu/pkg/api"
)

// ExtractTextFromPDF extracts text from a PDF file
func ExtractTextFromPDF(data []byte) (string, error) {
	var err error
	buf := bytes.NewReader(data)
	ctx, err := api.ReadContext(buf, nil)
	if err != nil {
		return "", err
	}

	var extractedText string
	for page := range ctx.PageCount {
		pIo, err := api.ExtractPage(ctx, page)
		if err != nil {
			return "", err
		}
		body := bytes.NewBuffer(nil)
		body.ReadFrom(pIo)
		extractedText += body.String()
	}

	return extractedText, nil
}
