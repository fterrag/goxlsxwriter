package goxlsxwriter

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include <xlsxwriter.h>
*/
import "C"
import "unsafe"

type FormatUnderline int

// Underline* constants are to be used with SetUnderline().
const (
	// Single underline.
	UnderlineSingle FormatUnderline = C.LXW_UNDERLINE_SINGLE

	// Double underline.
	UnderlineDouble FormatUnderline = C.LXW_UNDERLINE_DOUBLE

	// Single accounting underline.
	UnderlineSingleAccounting FormatUnderline = C.LXW_UNDERLINE_SINGLE_ACCOUNTING

	// Double accounting line.
	UndelineDoubleAccounting FormatUnderline = C.LXW_UNDERLINE_DOUBLE_ACCOUNTING
)

type FormatPattern int

// Pattern* constants are to be used with SetPattern().
const (
	// Solid.
	PatternSolid FormatPattern = C.LXW_PATTERN_SOLID

	// Medium gray.
	PatternMediumGray FormatPattern = C.LXW_PATTERN_MEDIUM_GRAY

	// Dark gray.
	PatternDarkGray FormatPattern = C.LXW_PATTERN_DARK_GRAY

	// Light gray.
	PatternLightGray FormatPattern = C.LXW_PATTERN_LIGHT_GRAY

	// Dark horizontal line.
	PatternDarkHorizontal FormatPattern = C.LXW_PATTERN_DARK_HORIZONTAL

	// Dark vertical line.
	PatternDarkVertical FormatPattern = C.LXW_PATTERN_DARK_VERTICAL

	// Dark diagonal stripe.
	PatternDarkDown FormatPattern = C.LXW_PATTERN_DARK_DOWN

	// Reverse dark diagonal stripe.
	PatternDarkUp FormatPattern = C.LXW_PATTERN_DARK_UP

	// Dark grid.
	PatternDarkGrid FormatPattern = C.LXW_PATTERN_DARK_GRID

	// Dark trellis.
	PatternDarkTrellis FormatPattern = C.LXW_PATTERN_DARK_TRELLIS

	// Light horizontal line.
	PatternLightHorizontal FormatPattern = C.LXW_PATTERN_LIGHT_HORIZONTAL

	// Light vertical line.
	PatternLightVertical FormatPattern = C.LXW_PATTERN_LIGHT_VERTICAL

	// Light diagonal stripe.
	PatternLightDown FormatPattern = C.LXW_PATTERN_LIGHT_DOWN

	// Reverse light diagonal stripe.
	PatternLightUp FormatPattern = C.LXW_PATTERN_LIGHT_UP

	// Light grid.
	PatternLightGrid FormatPattern = C.LXW_PATTERN_LIGHT_GRID

	// Light trellis.
	PatternLightTrellis FormatPattern = C.LXW_PATTERN_LIGHT_TRELLIS

	// 12.5% gray.
	PatternGray125 FormatPattern = C.LXW_PATTERN_GRAY_125

	// 6.25% gray.
	PatternGray625 FormatPattern = C.LXW_PATTERN_GRAY_0625
)

// Format represents an Excel style used to apply formatting to cells.
type Format struct {
	CFormat *C.struct_lxw_format
}

// NewFormat creates and returns a new instance of Format.
func NewFormat(workbook *Workbook) *Format {
	cFormat := C.workbook_add_format(workbook.CWorkbook)

	format := &Format{
		CFormat: cFormat,
	}

	return format
}

// SetFontName sets the format's font face.
func (f *Format) SetFontName(fontName string) {
	cFontName := C.CString(fontName)
	defer C.free(unsafe.Pointer(cFontName))

	C.format_set_font_name(f.CFormat, cFontName)
}

// SetFontSize sets the font size.
func (f *Format) SetFontSize(size int) {
	cSize := (C.double)(size)

	C.format_set_font_size(f.CFormat, cSize)
}

// SetFontColor sets the font color.
func (f *Format) SetFontColor(color int) {
	cColor := (C.lxw_color_t)(color)

	C.format_set_font_color(f.CFormat, cColor)
}

// SetBold sets the font to be bold.
func (f *Format) SetBold() {
	C.format_set_bold(f.CFormat)
}

// SetItalic sets the font to be italic.
func (f *Format) SetItalic() {
	C.format_set_italic(f.CFormat)
}

// SetUnderline sets the font to be underline using the
// specified UNDERLINE_* style.
func (f *Format) SetUnderline(style FormatUnderline) {
	cStyle := (C.uint8_t)(style)

	C.format_set_underline(f.CFormat, cStyle)
}

// SetPattern sets the pattern to the specific PATTERN_* pattern.
func (f *Format) SetPattern(pattern FormatPattern) {
	cPattern := (C.uint8_t)(pattern)

	C.format_set_pattern(f.CFormat, cPattern)
}

// SetBackgroundColor sets the background color.
func (f *Format) SetBackgroundColor(color int) {
	cColor := (C.lxw_color_t)(color)

	C.format_set_bg_color(f.CFormat, cColor)
}

// SetNumericalFormat sets the numerical format.
// It controls whether a number is displayed as an integer,
// a floating point number, a date, a currency value or some other
// user defined format (e.g., "d mmm yyyy").
func (f *Format) SetNumericalFormat(numberFormat string) {
	cNumberFormat := C.CString(numberFormat)
	defer C.free(unsafe.Pointer(cNumberFormat))

	C.format_set_num_format(f.CFormat, cNumberFormat)
}
