package main

import (
	"flag"
	"fmt"
	"encoding/csv"
	"os"
	"path/filepath"
)

func main() {
	var (
		f = flag.String("path", "", "CSV file path.")
		o = flag.String("output", "", "Output Base path.")
	)
	flag.Parse()

	file, err := os.Open(*f)
	if err != nil {
		panic(err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	var line []string

	for {
		line, err = reader.Read()
		if err != nil {
			break
		}
		divideRow(line, *o)
		fmt.Println(line[0])
	}
}

func divideRow(line []string, outputBase string) {
	if err := os.RemoveAll(filepath.FromSlash(outputBase + "/" + "_" + line[0])); err != nil {
		panic(err)
	}
	if err := os.Mkdir(filepath.FromSlash(outputBase + "/" + "_" + line[0]), 0777); err != nil {
		panic(err)
	}
	fmt.Println(line)

	var parent = line[0]
	file_batch, err := os.Create(filepath.FromSlash(
		fmt.Sprintf("%s/_%s/BATCH.csv", outputBase, parent)))
	if err != nil {
		panic(err)
	}
	defer file_batch.Close()

	writer := csv.NewWriter(file_batch)
	writer.Write([]string{line[1],line[0]})
	writer.Flush()

	var i = 0
	var parts = line[5:len(line)]

	file_id, err := os.Create(filepath.FromSlash(
		fmt.Sprintf("%s/_%s/ID.csv", outputBase, parent)))
	if err != nil {
		panic(err)
	}
	defer file_id.Close()
	writer_id := csv.NewWriter(file_id)
	for len(parts) > 0 {
		i += 1
		writer_id.Write([]string{parts[0],parts[0],parent})
		parts = parts[2:len(parts)]
	}
	writer_id.Flush()
}
