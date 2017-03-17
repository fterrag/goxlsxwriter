package xlsxwriter

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include <xlsxwriter.h>
*/
import "C"
import "unsafe"

const (
	UNDERLINE_SINGLE            int = C.LXW_UNDERLINE_SINGLE
	UNDERLINE_DOUBLE            int = C.LXW_UNDERLINE_DOUBLE
	UNDERLINE_SINGLE_ACCOUNTING int = C.LXW_UNDERLINE_SINGLE_ACCOUNTING
	UNDERLINE_DOUBLE_ACCOUNTING int = C.LXW_UNDERLINE_DOUBLE_ACCOUNTING
)
const (
	PATTERN_SOLID            int = C.LXW_PATTERN_SOLID
	PATTERN_MEDIUM_GRAY      int = C.LXW_PATTERN_MEDIUM_GRAY
	PATTERN_DARK_GRAY        int = C.LXW_PATTERN_DARK_GRAY
	PATTERN_LIGHT_GRAY       int = C.LXW_PATTERN_LIGHT_GRAY
	PATTERN_DARK_HORIZONTAL  int = C.LXW_PATTERN_DARK_HORIZONTAL
	PATTERN_DARK_VERTICAL    int = C.LXW_PATTERN_DARK_VERTICAL
	PATTERN_DARK_DOWN        int = C.LXW_PATTERN_DARK_DOWN
	PATTERN_DARK_UP          int = C.LXW_PATTERN_DARK_UP
	PATTERN_DARK_GRID        int = C.LXW_PATTERN_DARK_GRID
	PATTERN_DARK_TRELLIS     int = C.LXW_PATTERN_DARK_TRELLIS
	PATTERN_LIGHT_HORIZONTAL int = C.LXW_PATTERN_LIGHT_HORIZONTAL
	PATTERN_LIGHT_VERTICAL   int = C.LXW_PATTERN_LIGHT_VERTICAL
	PATTERN_LIGHT_DOWN       int = C.LXW_PATTERN_LIGHT_DOWN
	PATTERN_LIGHT_UP         int = C.LXW_PATTERN_LIGHT_UP
	PATTERN_LIGHT_GRID       int = C.LXW_PATTERN_LIGHT_GRID
	PATTERN_LIGHT_TRELLIS    int = C.LXW_PATTERN_LIGHT_TRELLIS
	PATTERN_GRAY_125         int = C.LXW_PATTERN_GRAY_125
	PATTERN_GRAY_0625        int = C.LXW_PATTERN_GRAY_0625
)

type Format struct {
	CFormat *C.struct_lxw_format
}

func NewFormat(workbook *Workbook) *Format {
	cFormat := C.workbook_add_format(workbook.CWorkbook)

	format := &Format{
		CFormat: cFormat,
	}

	return format
}

func (f *Format) SetFontName(fontName string) {
	cFontName := C.CString(fontName)
	defer C.free(unsafe.Pointer(cFontName))

	C.format_set_font_name(f.CFormat, cFontName)
}

func (f *Format) SetFontSize(size int) {
	cSize := (C.uint16_t)(size)

	C.format_set_font_size(f.CFormat, cSize)
}

func (f *Format) SetFontColor(color int) {
	cColor := (C.lxw_color_t)(color)

	C.format_set_font_color(f.CFormat, cColor)
}

func (f *Format) SetBold() {
	C.format_set_bold(f.CFormat)
}

func (f *Format) SetItalic() {
	C.format_set_italic(f.CFormat)
}

func (f *Format) SetUnderline(style int) {
	cStyle := (C.uint8_t)(style)

	C.format_set_underline(f.CFormat, cStyle)
}

func (f *Format) SetPattern(pattern int) {
	cPattern := (C.uint8_t)(pattern)

	C.format_set_pattern(f.CFormat, cPattern)
}

func (f *Format) SetBackgroundColor(color int) {
	cColor := (C.lxw_color_t)(color)

	C.format_set_bg_color(f.CFormat, cColor)
}

func (f *Format) SetNumberFormat(numberFormat string) {
	cNumberFormat := C.CString(numberFormat)
	defer C.free(unsafe.Pointer(cNumberFormat))

	C.format_set_num_format(f.CFormat, cNumberFormat)
}
