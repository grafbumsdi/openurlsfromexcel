package main

import (
	"flag"
	"github.com/tealeg/xlsx"
	"github.com/toqueteos/webbrowser"
	"io"
	"io/ioutil"
	"log"
	"math"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

func main() {
	// command line flags
	fileNamePtr := flag.String("filename", "https://www.dropbox.com/s/qv6qi3oejylsvrm/%C3%9Cbersicht%20Module.xlsx?dl=1", "can be a local file or downloadable file starting with 'http(s)://'")
	cellRangePtr := flag.String("cellrange", "L2:L20", "a valid xls cell range expression")
	flag.Parse()
	
	fileName := *fileNamePtr
	match, err := regexp.MatchString("http(s?)://", fileName)
	if(err != nil) {
		log.Fatalln("Error while parsing filename:", fileName)
	}
	if(match) {
		file, err := ioutil.TempFile(os.TempDir(), "tmpfile")
		if(err != nil) {
			log.Fatalln("Error while creating temp file")
		}
		defer os.Remove(file.Name())
		downloadFromUrl(fileName, file.Name())
		fileName = file.Name()
	} else {
		log.Println("Trying to open local file", fileName)
	}
	excelFile, err := filepath.Abs(fileName)
	if(err != nil) {
		log.Fatalln("Error while parsing", fileName)
	}
	getUrlsFromExcelCellRange(excelFile, *cellRangePtr, 0)
}

func getUrlsFromExcelCellRange(excelFileName string, cellRange string, sheetIndex int) {
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		log.Fatalln("Error while opening excel file", excelFileName)
	}
	// sheet := xlFile.Sheets[sheetIndex]
	sheet := xlFile.Sheets[0]
	log.Println("Trying to parse cell range", cellRange)
	columnStart, columnEnd, rowStart, rowEnd := parseRange(cellRange)
	for r := rowStart - 1; r < rowEnd; r++ {
		row := sheet.Rows[r]
		for c := columnStart - 1; c < columnEnd; c++ {
			cellValue := row.Cells[c].String()
			match, err := regexp.MatchString("http(s?)://", cellValue)
			if(err != nil) {
				log.Println("Error while trying to parse value", cellValue, "from cell rownum", r, "colnum", c)
			}
			if(match) {
				openUrl(cellValue)
			}
		}
	}
}

func openUrl(url string) {
	log.Println("Trying to open the following url:", url)
	webbrowser.Open(url)
}

func parseRange(cellRange string) (columnStart, columnEnd, rowStart, rowEnd int){
	r := regexp.MustCompile("([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)")
	res := r.FindStringSubmatch(cellRange)
	// first result is always the whole regex pattern (e.g: A12:B32)
	// all further results are sub 'patterns'/'groups' that were found (e.g: A, 12, B, 32)
	if res[0] != cellRange || len(res) != 5 {
		log.Fatalln("The given cell range", cellRange, "was not valid")
	}
	a, _ := strconv.Atoi(res[2])
	b, _ := strconv.Atoi(res[4])
	rowStart, rowEnd = orderAsc(a, b)
	columnStart, columnEnd = orderAsc(convertStringToColumnIndex(res[1]), convertStringToColumnIndex(res[3]))
	return
}

func orderAsc(a int, b int) (int, int) {
	if(a > b) {
		return b, a
	}
	return a, b
}

// columnindex starting with "A" = 1
// for example: "AB" = 28
func convertStringToColumnIndex(columnName string) int{
	var val float64
	cl := strings.ToLower(columnName)
	poweredBy := len(cl)
	for index, c := range cl {
		charValue := c - 96 // - "a"[0] not working for some reason
		val += float64(charValue) * (math.Pow(26, float64(poweredBy - index - 1)))
	}
	return int(val)
}

func downloadFromUrl(url string, fileName string) {
	log.Println("Downloading", url, "to", fileName)

	// TODO: check file existence first with io.IsExist
	output, err := os.Create(fileName)
	if err != nil {
		log.Println("Error while creating", fileName, "-", err)
		return
	}
	defer output.Close()

	response, err := http.Get(url)
	if err != nil {
		log.Println("Error while downloading", url, "-", err)
		return
	}
	defer response.Body.Close()

	n, err := io.Copy(output, response.Body)
	if err != nil {
		log.Println("Error while downloading", url, "-", err)
		return
	}

	log.Println(n, "bytes downloaded.")
}