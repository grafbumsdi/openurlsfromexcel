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
	
	// if given filename was an URL we have to download it into a temp file
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
	
	// open excelfile and open URLs from it
	excelFile, err := filepath.Abs(fileName)
	if(err != nil) {
		log.Fatalln("Error while parsing", fileName)
	}
	openUrlsFromExcelCellRange(excelFile, *cellRangePtr, 0)
}

// Opens the given excel file, reads the value of each cell in the given cell range
// and tries to open each value in the default browser
func openUrlsFromExcelCellRange(excelFileName string, cellRange string, sheetIndex int) {
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		log.Fatalln("Error while opening excel file", excelFileName)
	}
	log.Println("Excel File has", len(xlFile.Sheets), "sheets")
	sheet := xlFile.Sheets[0]
	log.Println("First Sheet has", len(sheet.Rows), "rows")
	
	log.Println("Trying to parse cell range", cellRange)
	columnStart, columnEnd, rowStart, rowEnd := parseRange(cellRange)
	log.Println("Parsed to (zero-based) rowStart:", rowStart - 1, "rowEnd:", rowEnd - 1, "columnStart:", columnStart - 1, "columnEnd:", columnEnd - 1)
	
	for r := rowStart - 1; r < rowEnd; r++ {
		if len(sheet.Rows) <= r  {
			log.Println("There is no row with rownum", r)
			continue
		}
		row := sheet.Rows[r]
		for c := columnStart - 1; c < columnEnd; c++ {
			if (row == nil || row.Cells == nil || len(row.Cells) <= c) {
				log.Println("Does NOT contain a valid cell value in rownum", r, "colnum", c)
				continue
			}
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

// Opens the given Url with the default browser of the user
func openUrl(url string) {
	log.Println("Trying to open the following url:", url)
	webbrowser.Open(url)
}

// Parses the given cell range string (e.g: B12:C3)
// Returns the 4 coordinates of the range: column start/end and row start/end starting from index 1
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

// Returns the given values in ascending order
func orderAsc(a int, b int) (int, int) {
	if(a > b) {
		return b, a
	}
	return a, b
}

// Converts a given column identifier in string format (e.g: "AB") to the numeric column index
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

// Downloads the response body of the given URL into the given file
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