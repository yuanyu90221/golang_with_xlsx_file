package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	categories := map[string]string{"A2": "Small", "A3": "Normal", "A4": "Large", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}
	f := excelize.NewFile()
	// f.SetCellValue("Sheet1", "A1", "test")

	f.NewSheet("FirstSheet")
	for k, v := range categories {
		f.SetCellValue("FirstSheet", k, v)
	}
	for k, v := range values {
		f.SetCellValue("FirstSheet", k, v)
	}
	if err := f.AddChart("FirstSheet", "E1", `{"type":"col3DClustered","series":[{"name":"FirstSheet!$A$2","categories":"FirstSheet!$B$1:$D$1","values":"FirstSheet!$B$2:$D$2"},{"name":"FirstSheet!$A$3","categories":"FirstSheet!$B$1:$D$1","values":"FirstSheet!$B$3:$D$3"},{"name":"FirstSheet!$A$4","categories":"FirstSheet!$B$1:$D$1","values":"FirstSheet!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`); err != nil {
		fmt.Println(err)
		return
	}
	f.DeleteSheet("Sheet1")
	if err := f.SaveAs("my.xlsx"); err != nil {
		fmt.Println(err)
	}
	fmt.Println("my.xlsx created")
}
