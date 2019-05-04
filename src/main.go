package main

import (
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"github.com/mitchellh/reflectwalk"
	"github.com/tealeg/xlsx"
	"os"
	"reflect"
	"time"
)

type Sheet struct {
	Name string `json:"name"`
	//  File        *File
	Rows     []*Row `json:"rows"`
	Cols     []*Col `json:"cols"`
	MaxRow   int    `json:"maxrow"`
	MaxCol   int    `json:"maxcol"`
	Hidden   bool   `json:"hidden"`
	Selected bool   `json:"selected"`
	//  SheetViews  []SheetView
	//  SheetFormat SheetFormat
	//  AutoFilter  *AutoFilter
	ErrorCells    []*Cell `json:"errorcells"`
	OrgCellCount  int     `json:"orgcellcount"`
	CellCount     int     `json:"cellcount"`
	RowCount      int     `json:"rowcount"`
	DropCellCount int     `json:"dropcellcount"`
	Generator     string  `json:"generator"`
	Date          string  `json:"date"`
}

type Row struct {
	Cells  []*Cell `json:"cells"`
	Hidden bool    `json:"hidden"`
	//  Sheet        *Sheet
	Height float64 `json:"height"`
	//  OutlineLevel uint8
}

type Col struct {
	Min       int     `json:"min"`
	Max       int     `json:"max"`
	Hidden    bool    `json:"hidden"`
	Width     float64 `json:"width"`
	Collapsed bool    `json:"collapsed"`
	//  OutlineLevel uint8   `json:"outlinelevel"`
	//  DataValidation []*xlsxCellDataValidation
}

type Cell struct {
	//  Row   *Row
	Id     string `json:"id"`
	X      int    `json:"X"`
	Y      int    `json:"Y"`
	Value  string `json:"value"`
	NumFmt string `json:"numfmt"`
	Hidden bool   `json:"hidden"`
	HMerge int    `json:"hmerge"`
	VMerge int    `json:"vmerge"`
	Style  Style  `json:"style"`
	Error  string `json:"error"`
	//  DataValidation *xlsxCellDataValidation
}

type Style struct {
	Border         Border    `json:"border"`
	Fill           Fill      `json:"fill"`
	Font           Font      `json:"font"`
	ApplyBorder    bool      `json:"applyborder"`
	ApplyFill      bool      `json:"applyfill"`
	ApplyFont      bool      `json:"applyfont"`
	ApplyAlignment bool      `json:"applyalignment"`
	Alignment      Alignment `json:"alignment"`
	//   NamedStyleIndex *int
}

type Border struct {
	Left        string `json:"left"`
	LeftColor   string `json:"leftcolor"`
	Right       string `json:"right"`
	RightColor  string `json:"rightcolor"`
	Top         string `json:"top"`
	TopColor    string `json:"topcolor"`
	Bottom      string `json:"bottom"`
	BottomColor string `json:"bottomcolor"`
}

type Fill struct {
	PatternType string `json:"patterntype"`
	BgColor     string `json:"bgcolor"`
	FgColor     string `json:"fgcolor"`
}

type Font struct {
	Size      int    `json:"size"`
	Name      string `json:"name"`
	Family    int    `json:"family"`
	Charset   int    `json:"charset"`
	Color     string `json:"color"`
	Bold      bool   `json:"bold"`
	Italic    bool   `json:"italic"`
	Underline bool   `json:"underline"`
}

type Alignment struct {
	Horizontal   string `json:"horizontal"`
	Indent       int    `json:"indent"`
	ShrinkToFit  bool   `json:"shrinktofit"`
	TextRotation int    `json:"textrotation"`
	Vertical     string `json:"vertical"`
	WrapText     bool   `json:"wraptext"`
}


func usage() {
	fmt.Fprintf(os.Stderr, `usage of %s:
   %s [options] FILENAME SHEETNAME

 [options]`,
		os.Args[0], os.Args[0])

	fmt.Println("")
	flag.PrintDefaults()
	fmt.Println("")

}

func main() {

	flag.Usage = usage

	var (
		minmam          = flag.Bool("m", false, "jsondata min size")
		output_filename = flag.String("o", "stdout", "output filename option")
		format          = flag.Bool("f", false, "json format option")
		reflect         = flag.Bool("r", false, "reflectwalk option")
	)

	flag.Parse()

	args := flag.Args()

	if len(args) != 2 {
		fmt.Println("*** argument less")
		flag.Usage()
		return
	}

	//fmt.Println("output:", *output_filename)
	//fmt.Println("format:", *format)
	//fmt.Println("arg1:", args[0])
	//fmt.Println("arg1:", args[1])

	input_filename := args[0]
	input_sheetname := args[1]

	if _, err := os.Stat(input_filename); os.IsNotExist(err) {
		fmt.Println("*** File does not exist:", input_filename)
		return
	}
	//excelFileName := "Book1.xlsx"
	excelFileName := input_filename
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return
	}
	//for _, sheet := range xlFile.Sheets {
	//sheet := xlFile.Sheet["Sheet1"]
	_, exist := xlFile.Sheet[input_sheetname]

	if !exist {
		fmt.Println("*** sheet not exist:", input_sheetname)
		return
	}

	sheet := xlFile.Sheet[input_sheetname]
	_sheet := Sheet{}
	_sheet.Name = sheet.Name
	_sheet.MaxRow = sheet.MaxRow
	_sheet.MaxCol = sheet.MaxCol
	_sheet.Hidden = sheet.Hidden
	_sheet.Selected = sheet.Selected

	_sheet.Rows = make([]*Row, 0)

	_sheet.Cols = make([]*Col, 0)
	_y := 0

	_org_cell_count := 0
	_cell_count := 0
	_row_count := 0
	_drop_cell_count := 0
	for _, row := range sheet.Rows {
		_row := Row{}
		_row.Hidden = row.Hidden
		_row.Height = row.Height
		_row.Cells = make([]*Cell, 0)

		_sheet.Rows = append(_sheet.Rows, &_row)

		_x := 0
		for _, cell := range row.Cells {
			var _xlsxStyle *xlsx.Style
			//var _xlsxBorder  *xlsx.Border
			//var _xlsxFill  *xlsx.Fill
			//var _xlsxFont  *xlsx.Font
			//var _xlsxAlignment  *xlsx.Alignment

			_cell := Cell{}
			_cell.Value = cell.Value
			_cell.NumFmt = cell.NumFmt
			_cell.Hidden = cell.Hidden
			_cell.HMerge = cell.HMerge
			_cell.VMerge = cell.VMerge

			_xlsxStyle = cell.GetStyle()
			_cell.Style = Style{}
			_cell.Style.ApplyBorder = _xlsxStyle.ApplyBorder
			_cell.Style.ApplyFill = _xlsxStyle.ApplyFill
			_cell.Style.ApplyFont = _xlsxStyle.ApplyFont
			_cell.Style.ApplyAlignment = _xlsxStyle.ApplyAlignment

			_cell.Style.Border = Border{}
			_cell.Style.Border.Left = _xlsxStyle.Border.Left
			_cell.Style.Border.LeftColor = _xlsxStyle.Border.LeftColor
			_cell.Style.Border.Right = _xlsxStyle.Border.Right
			_cell.Style.Border.RightColor = _xlsxStyle.Border.RightColor
			_cell.Style.Border.Top = _xlsxStyle.Border.Top
			_cell.Style.Border.TopColor = _xlsxStyle.Border.TopColor
			_cell.Style.Border.Bottom = _xlsxStyle.Border.Bottom
			_cell.Style.Border.BottomColor = _xlsxStyle.Border.BottomColor

			_cell.Style.Fill = Fill{}
			_cell.Style.Fill.PatternType = _xlsxStyle.Fill.PatternType
			_cell.Style.Fill.BgColor = _xlsxStyle.Fill.BgColor
			_cell.Style.Fill.FgColor = _xlsxStyle.Fill.FgColor

			_cell.Style.Font = Font{}
			_cell.Style.Font.Size = _xlsxStyle.Font.Size
			_cell.Style.Font.Name = _xlsxStyle.Font.Name
			_cell.Style.Font.Family = _xlsxStyle.Font.Family
			_cell.Style.Font.Charset = _xlsxStyle.Font.Charset
			_cell.Style.Font.Color = _xlsxStyle.Font.Color
			_cell.Style.Font.Bold = _xlsxStyle.Font.Bold
			_cell.Style.Font.Italic = _xlsxStyle.Font.Italic
			_cell.Style.Font.Underline = _xlsxStyle.Font.Underline

			_cell.Style.Alignment = Alignment{}
			_cell.Style.Alignment.Horizontal = _xlsxStyle.Alignment.Horizontal
			_cell.Style.Alignment.Indent = _xlsxStyle.Alignment.Indent
			_cell.Style.Alignment.ShrinkToFit = _xlsxStyle.Alignment.ShrinkToFit
			_cell.Style.Alignment.TextRotation = _xlsxStyle.Alignment.TextRotation
			_cell.Style.Alignment.Vertical = _xlsxStyle.Alignment.Vertical
			_cell.Style.Alignment.WrapText = _xlsxStyle.Alignment.WrapText

			_cell.Id = xlsx.GetCellIDStringFromCoords(_x, _y)
			_cell.X = _x + 1
			_cell.Y = _y + 1

			// vacant cell judgement

			celljudg := true

			if _cell.Value == "" &&
				_cell.HMerge == 0 &&
				_cell.VMerge == 0 &&
				(_cell.Style.Border.Left == "" || _cell.Style.Border.Left == "none") &&
				(_cell.Style.Border.Right == "" || _cell.Style.Border.Right == "none") &&
				(_cell.Style.Border.Top == "" || _cell.Style.Border.Top == "none") &&
				(_cell.Style.Border.Bottom == "" || _cell.Style.Border.Bottom == "none") &&
				(_cell.Style.Fill.PatternType == "" || _cell.Style.Fill.PatternType == "none") &&
				(_cell.Style.Fill.BgColor == "" || _cell.Style.Fill.BgColor == "none") &&
				(_cell.Style.Fill.FgColor == "" || _cell.Style.Fill.FgColor == "none") &&
				true {
				celljudg = false
			}

			if *minmam {
				if celljudg {
					_row.Cells = append(_row.Cells, &_cell)
					_cell_count++
				} else {
					_drop_cell_count++
				}

			} else {

				_row.Cells = append(_row.Cells, &_cell)
				_cell_count++

			}

			_x++
			_org_cell_count++
		}
		_y++
		_row_count++
	}
	for _, col := range sheet.Cols {
		_col := Col{}
		_col.Hidden = col.Hidden
		_col.Width = col.Width
		_col.Min = col.Min
		_col.Max = col.Max
		_col.Hidden = col.Hidden
		_col.Collapsed = col.Collapsed

		_sheet.Cols = append(_sheet.Cols, &_col)
	}
	_sheet.OrgCellCount = _org_cell_count
	_sheet.CellCount = _cell_count
	_sheet.RowCount = _row_count
	_sheet.DropCellCount = _drop_cell_count
	_sheet.Generator = "xlsx2json"
	_sheet.Date = time.Now().String()

	/******************************************************/
	//var walker interface{}
	if *reflect {
		var walker Walker
		if err := reflectwalk.Walk(_sheet, walker); err != nil {
			panic(err)
		}
	}

	/******************************************************/

	jsonBytes, err := json.Marshal(_sheet)

	if err != nil {
		//log.Fatal(err)
		fmt.Println(err)
	}

	if json.Valid(jsonBytes) {
		fmt.Println("+++ marshal Valid ..")
	} else {
		fmt.Println("*** marshal inValid ..")
	}

	if *output_filename == "stdout" {
		if *format {
			//fmt.Println(string(jsonBytes))
			out := new(bytes.Buffer)
			json.Indent(out, jsonBytes, "", "    ")
			fmt.Println(out.String())

		} else {
			fmt.Println(string(jsonBytes))
		}
	} else {
		if *format {
			//fmt.Println(string(jsonBytes))
			out := new(bytes.Buffer)
			json.Indent(out, jsonBytes, "", "    ")
			//fmt.Println(out.String())
			file, err := os.Create(*output_filename)
			check(err)
			defer file.Close()

			//fmt.Fprintf(file, out.String())
			_, err = file.Write(out.Bytes())
			check(err)

		} else {
			//fmt.Println(string(jsonBytes))
			file, err := os.Create(*output_filename)
			check(err)
			defer file.Close()
			//fmt.Fprintf(file, string(jsonBytes))
			_, err = file.Write(jsonBytes)
			check(err)
		}
	}

	fmt.Println("org cell count:", _org_cell_count)
	fmt.Println("cell count:", _cell_count)
	fmt.Println("row count:", _row_count)
	fmt.Println("drop cell count:", _drop_cell_count)
	//}
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}

/*

type PrimitiveWalker interface {
    Primitive(reflect.Value) error
}

*/

type Walker struct {
	base int
}

func (w Walker) Enter(loc reflectwalk.Location) error {
	fmt.Print("------")
	fmt.Println(loc)
	return nil
}
func (w Walker) Struct(v reflect.Value) error {
	fmt.Println("struct: ")
	//fmt.Println( v )
	return nil
}
func (w Walker) StructField(f reflect.StructField, v reflect.Value) error {
	fmt.Print("field: ")
	//fmt.Print( f )
	fmt.Println(v)
	return nil
}
func (w Walker) Primitive(v reflect.Value) error {
	w.base++
	//fmt.Println( "   :" + v.String())
	fmt.Println(v)
	return nil
}
