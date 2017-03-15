package xlsxwriter

import (
	"encoding/hex"
	"fmt"
	"math/rand"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
)

func MakeTestWorkbook() (*Workbook, string) {
	filename := tempFileName("xlsxwriter", ".xlsx")
	workbook := NewWorkbook(filename, nil)

	return workbook, filename
}

func tempFileName(prefix string, suffix string) string {
	randBytes := make([]byte, 16)
	rand.Read(randBytes)

	return filepath.Join(os.TempDir(), prefix+hex.EncodeToString(randBytes)+suffix)
}

func CompareXlsxFiles(t *testing.T, expectedPath string, generatedPath string) {
	code := fmt.Sprintf("import helper_functions; print helper_functions._compare_xlsx_files('%s', '%s', [], [])", generatedPath, expectedPath)

	cmd := exec.Command("python", "-c", code)
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatal(err)
	}

	if string(out) != "('Ok', 'Ok')\n" {
		t.Fatalf("%s and %s are not identical", expectedPath, generatedPath)
	}
}
