package goxlsxwriter

import (
	"fmt"
	"io/ioutil"
	"os/exec"
	"testing"
)

func MakeTestWorkbook() (*Workbook, string) {
	filename := tempFile("goxlsxwriter")
	workbook := NewWorkbook(filename, nil)

	return workbook, filename
}

func tempFile(prefix string) string {
	tempFile, err := ioutil.TempFile("", prefix)
	if err != nil {
		panic(err)
	}

	return tempFile.Name()
}

func CompareXlsxFiles(t *testing.T, expectedPath string, generatedPath string) {
	code := fmt.Sprintf("import sys; sys.path.append('./resources'); import helper_functions; print helper_functions._compare_xlsx_files('%s', '%s', [], [])", generatedPath, expectedPath)

	cmd := exec.Command("python", "-B", "-c", code)
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatal(err)
	}

	outStr := string(out)

	if outStr != "('Ok', 'Ok')\n" {
		t.Fatalf("%s and %s are not identical: %s", expectedPath, generatedPath, outStr)
	}
}
