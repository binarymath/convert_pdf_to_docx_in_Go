package main

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/unidoc/unipdf/v3/common"
	"github.com/unidoc/unipdf/v3/processor"
)

func main() {

	f, err := os.Open("arquivo.pdf")
	if err != nil {
		panic(err)
	}
	defer f.Close()


	pdfProcessor, err := processor.New()
	if err != nil {
		panic(err)
	}

	err = pdfProcessor.Load(f)
	if err != nil {
		panic(err)
	}


	options := processor.ConvertToTextOptions{
		ExtractAnnotations: true,
		Flatten:            true,
	}


	text, err := pdfProcessor.ConvertToText(&options)
	if err != nil {
		panic(err)
	}

	file := excelize.NewFile()
	sheet := file.NewSheet("Sheet1")


	err = file.SetCellValue("Sheet1", "A1", text)
	if err != nil {
		panic(err)
	}

	
	err = file.SaveAs("arquivo.docx")
	if err != nil {
		panic(err)
	}

	fmt.Println("PDF convertido para arquivo do Word com sucesso.")
}
