package main

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"regexp"
	"strings"
	"sync"

	// "time"

	// "sync"

	"github.com/xuri/excelize/v2"
)

var start = 310
var end   = 316
var name  = fmt.Sprintf("%d-%d", start, end)

func execute(row []string, i int) error {
	downloadURL := row[1]
	if strings.Contains(row[1], "https://drive.google.com") {
		// return nil
		re := regexp.MustCompile(`https://drive.google.com/file/d/([a-zA-Z0-9_-]+)/view`)
		match := re.FindStringSubmatch(row[1])

		if len(match) < 2 {
			return fmt.Errorf("invalid Google Drive URL on row : %d", i)
		}

		// Extracted file ID
		fileID := match[1]
		// fmt.Println("Extracted File ID:", fileID)

		// Construct the direct download URL
		downloadURL = fmt.Sprintf("https://drive.google.com/uc?export=download&id=%s", fileID)
		// fmt.Println("Download URL:", downloadURL)
	}

	outFile, err := os.Create(fmt.Sprintf("%s/%s.pdf", name, row[0]))
	if err != nil {
		fmt.Println("Error creating file:", err)
		return err
	}
	defer outFile.Close()

	// Download the file
	response, err := http.Get(downloadURL)
	if err != nil {
		fmt.Println("Error downloading file row :", i)
		return err
	}
	defer response.Body.Close()

	// Write the response body to the file
	_, err = io.Copy(outFile, response.Body)
	if err != nil {
		fmt.Println("Error saving file:", err)
		return err
	}

	return nil
}

func main() {
	f, err := excelize.OpenFile("./test-cvt.xlsx")
	if err != nil {
		fmt.Println(err.Error())
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err.Error())
		}
	}()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err.Error())
	}

	os.MkdirAll(name, 0755)

	var wg sync.WaitGroup
	var mu sync.Mutex
	hehe := []error{}
	for i := range rows {
		if i >= start && i <= end {
			if len(rows[i]) >= 2 {
				wg.Add(1)
				go func() {
					if rows[i][1] != "link" {
						e := execute(rows[i], i)
						fmt.Println("Row : ", i)
						
						if e != nil {
							mu.Lock()
							defer mu.Unlock()
							hehe = append(hehe, e)
						}
					}
					wg.Done()
				}()
			}
		}
	}

	wg.Wait()

	for _, err := range hehe {
		if err != nil {
			fmt.Println("Error:", err)
		}
	}

	fmt.Println("OUTTT")
}
