package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	//  Открываем необходимые файлы (Ласточка, льготники и копию льготников(надо ли?))
	lastochka_file, err := excelize.OpenFile("Ласточка 2021.xlsm")
	if err != nil {
		log.Fatal(err)
	}

	// szs_file, err := excelize.OpenFile("SZS1.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	// rec_file, err := excelize.OpenFile("SZS_rec.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	//  Создаем массив с ячейками в нашей таблице, в которой хранятся тарифы
	tariff_cells := [6]string{
		"D2504",
		"D2505",
		"D2506",
		"D2507",
		"D2509",
		"B2499",
	}

	//  Создаем массив с названиями тарифов
	tariff_names := [6]string{
		"ОДН на ХВС",
		"ОДН на ГВС",
		"ОДН на электро",
		"ОДН на водоотв",
		"Содержание",
		"Электроэнергия",
	}

	
	type Flat struct {
		number int
		owner string
		area float64
		power int
	}

	type House struct {
		flat []Flat
	}

	//  Получаем значения тарифов и формируем карту [Тариф: значение]
	tariffs := make(map[string]float64)

	for idx, el := range tariff_cells {
		cell, err := lastochka_file.GetCellValue("Квитанции_чистые", el)
		if err != nil {
			log.Fatal(err)
		}
		if cellfloat, err := strconv.ParseFloat(cell, 64); err == nil {
			tariffs[tariff_names[idx]] = cellfloat
		}
	}

	// Для проверки печатаем тарифы на экран
	for idx, el := range tariffs {
		fmt.Println(idx, "\t", el)
	}

	//-----------------------------------------------------------------------------------------
	// // Формируем карту нужных полей в выходном документе
	// out_col := map[string]string{
	// 	"FACTP":   "L",
	// 	"TARIF":   "N",
	// 	"PRIZN":   "P",
	// 	"FACTOP":  "V",
	// 	"FACTOP2": "W",
	// }

	// Перебираем выходной документ построчно
	rows, err := lastochka_file.GetRows("Жильцы")
	if err != nil {
		fmt.Println(err)
		return
	}

	var CurrentFlat Flat

	// Перебираем строки, заносим в структуру и её в структуру общую
	for _, rows := range rows {
		CurrentFlat.number, _ = strconv.Atoi(rows[0])
		CurrentFlat.owner = rows[1] + " " + rows[2] + " " + rows[3]
		CurrentFlat.area, _ = strconv.ParseFloat(rows[4], 64)
		fmt.Println(CurrentFlat)
	}
	//
	// // Разбираем каждую строку и вносим значения тарифов в выходную таблицу
	// for idx, row := range rows {
	// 	cell_N := (out_col["TARIF"] + strconv.Itoa(idx+1))
	// 	if row[7] == "ОДН на ХВС" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["ОДН на ХВС"])
	// 	} else if row[7] == "ОДН на ГВС" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["ОДН на ГВС"])
	// 	} else if row[7] == "ОДН на водоотведение" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["ОДН на водоотв"])
	// 	} else if row[7] == "Электрическая энергия на общедомовые нужды" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["ОДН на электро"])
	// 	} else if row[7] == "Содержание жилья" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["Содержание"])
	// 	} else if row[7] == "Э: МЖД с ЦГВС и электроплитами" {
	// 		rec_file.SetCellValue("Лист1", cell_N, tariffs["Электроэнергия"])
	// 	}
	// }
	//
	// //  Сохраняем выходной файл
	// if err := rec_file.SaveAs("Book1.xlsx"); err != nil {
	// 	fmt.Println(err)
	// }

}
