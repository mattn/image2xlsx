package main

import (
	"flag"
	"fmt"
	"image"
	"image/draw"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"log"
	"os"
	"path/filepath"

	"github.com/tealeg/xlsx"
)

var (
	pixel = flag.Float64("p", 0.2, "pixel size")
)

func image2xlsx(name string) error {
	f, err := os.Open(name)
	if err != nil {
		return err
	}
	defer f.Close()
	img, _, err := image.Decode(f)
	if err != nil {
		return err
	}
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")
	if err != nil {
		return err
	}
	nrgba := image.NewNRGBA(img.Bounds())
	draw.Draw(nrgba, img.Bounds(), img, image.Pt(0, 0), draw.Src)
	for y := 0; y < img.Bounds().Dy(); y++ {
		row := sheet.AddRow()
		for x := 0; x < img.Bounds().Dx(); x++ {
			cell := row.AddCell()
			r, g, b, a := nrgba.At(x, y).RGBA()
			if a > 0 {
				c := fmt.Sprintf("%02X%02X%02X%02X", a&0xff, r&0xff, g&0xff, b&0xff)
				cell.GetStyle().ApplyFill = true
				cell.GetStyle().Fill = *xlsx.NewFill("solid", c, c)
			}
		}
	}
	sheet.SetColWidth(0, img.Bounds().Dx()-1, *pixel)
	for _, row := range sheet.Rows {
		row.SetHeightCM(*pixel * 0.0353 * 10)
	}

	name = name[:len(name)-len(filepath.Ext(name))] + ".xlsx"
	return file.Save(name)
}

func main() {
	flag.Parse()
	for _, arg := range flag.Args() {
		if err := image2xlsx(arg); err != nil {
			log.Fatal(err)
		}
	}
}
