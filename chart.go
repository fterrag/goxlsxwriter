package goxlsxwriter

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include <xlsxwriter.h>
*/
import "C"
import (
	"unsafe"
)

type ChartType int

const (
	// Area chart.
	chartTypeArea ChartType = C.LXW_CHART_AREA

	// Area chart - stacked.
	chartTypeAreaStacked ChartType = C.LXW_CHART_AREA_STACKED

	// Area chart - percentage stacked.
	chartTypeAreaStackedPercent ChartType = C.LXW_CHART_AREA_STACKED_PERCENT

	// Bar chart.
	chartTypeBar ChartType = C.LXW_CHART_BAR

	// Bar chart - stacked.
	chartTypeBarStacked ChartType = C.LXW_CHART_BAR_STACKED

	// Bar chart - percentage stacked.
	chartTypeBarStackedPercent ChartType = C.LXW_CHART_BAR_STACKED_PERCENT

	// Column chart.
	chartTypeColumn ChartType = C.LXW_CHART_COLUMN

	// Column chart - stacked.
	chartTypeColumnStacked ChartType = C.LXW_CHART_COLUMN_STACKED

	// Column chart - percentage stacked.
	chartTypeColumnStackedPercent ChartType = C.LXW_CHART_COLUMN_STACKED_PERCENT

	// Doughnut chart.
	chartTypeDoughnut ChartType = C.LXW_CHART_DOUGHNUT

	// Line chart.
	chartTypeLine ChartType = C.LXW_CHART_LINE

	// Pie chart.
	chartTypePie ChartType = C.LXW_CHART_PIE

	// Scatter chart.
	chartTypeScatter ChartType = C.LXW_CHART_SCATTER

	// Scatter chart - straight.
	chartTypeScatterStraight ChartType = C.LXW_CHART_SCATTER_STRAIGHT

	// Scatter chart - straight with markers.
	chartTypeScatterStraightWithMarkers ChartType = C.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS

	// Scatter chart - smooth.
	chartTypeScatterSmooth ChartType = C.LXW_CHART_SCATTER_SMOOTH

	// Scatter chart - smooth with markers.
	chartTypeScatterSmoothWithMarkers ChartType = C.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS

	// Radar chart.
	chartTypeRadar ChartType = C.LXW_CHART_RADAR

	// Radar chart - with markers.
	chartTypeRadarWithMarkers ChartType = C.LXW_CHART_RADAR_WITH_MARKERS

	// Radar chart - filled.
	chartTypeRadarFilled ChartType = C.LXW_CHART_RADAR_FILLED
)

// Chart represents an Excel chart.
type Chart struct {
	CChart *C.lxw_chart
}

// ChartSeries represents an Excel chart series.
type ChartSeries struct {
	Chart        *Chart
	CChartSeries *C.lxw_chart_series
}

// ChartOptions contains options to be set when inserting a chart into a
// worksheet.
type ChartOptions struct {
	ImageOptions
}

// NewChart creates and returns a new Chart.
func NewChart(workbook *Workbook, chartType ChartType) *Chart {
	cChartType := C.uchar(chartType)

	cChart := C.workbook_add_chart(workbook.CWorkbook, cChartType)

	chart := &Chart{
		CChart: cChart,
	}

	return chart
}

// NewChartSeries creates and returns a new ChartSeries.
func NewChartSeries(chart *Chart, categories string, values string) *ChartSeries {
	cCategories := C.CString(categories)
	defer C.free(unsafe.Pointer(cCategories))

	cValues := C.CString(values)
	defer C.free(unsafe.Pointer(cValues))

	cChartSeries := C.chart_add_series(chart.CChart, cCategories, cValues)

	chartSeries := &ChartSeries{
		Chart:        chart,
		CChartSeries: cChartSeries,
	}

	return chartSeries
}
