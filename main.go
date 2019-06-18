package main

import (
	"encoding/csv"
	"flag"
	"log"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	inFile := flag.String("input", "", "The input file (xlsx)")
	outFile := flag.String("output", "", "The output CSV")
	flag.Parse()

	if *inFile == "" || *outFile == "" {
		panic("Both infile and outfile parameters are required")
	}
	dumpData(*inFile, *outFile)
}

func dumpData(excelFilename string, outputFile string) {
	csvfile, err := os.Create(outputFile)
	if err != nil {
		log.Fatalf("failed creating file: %s", err)
	}
	defer csvfile.Close()

	xlFile, err := xlsx.OpenFile(excelFilename)
	if err != nil {
		panic("Failed to open file")
	}
	sheet := xlFile.Sheets[0]

	csvwriter := csv.NewWriter(csvfile)
	writeData(sheet, csvwriter)
}

func writeData(sheet *xlsx.Sheet, csvwriter *csv.Writer) {
	for _, row := range sheet.Rows {
		rowLength := len(row.Cells)
		if rowLength == 0 {
			break
		}

		idx := 0
		r := make([]string, rowLength)
		for _, cell := range row.Cells {
			r[idx] = strings.TrimSpace(cell.String())
			idx++
		}

		_ = csvwriter.Write(r)

	}

	csvwriter.Flush()
}
