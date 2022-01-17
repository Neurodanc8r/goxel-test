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

	var month int
	var month_cell int

	fmt.Println("Введите номер месяца, на который делаем расчет:")
	fmt.Println("-----------------")
	fmt.Println("1  - Январь")
	fmt.Println("2  - Февраль")
	fmt.Println("3  - Март")
	fmt.Println("4  - Апрель")
	fmt.Println("5  - Май")
	fmt.Println("6  - Июнь")
	fmt.Println("7  - Июль")
	fmt.Println("8  - Август")
	fmt.Println("9  - Сентябрь")
	fmt.Println("10 - Октябрь")
	fmt.Println("11 - Ноябрь")
	fmt.Println("12 - Декабрь")
	fmt.Println("-----------------")
	fmt.Print("::>")
	fmt.Scanln(&month)
	fmt.Println()

	month_index := [13]int {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}

	if month < 1 || month > 12 {
		fmt.Println("ERROR :: Неправильно введён месяц!")
		fmt.Println()
	} else {
		fmt.Println("OK")
		fmt.Println()
		month_cell = month_index[month]
	}

	tariff_cells := [6]string{
		"D2504",
		"D2505",
		"D2506",
		"D2507",
		"D2509",
		"B2499",
	}

	tariff_names := [6]string{
		"ОДН на ХВС",
		"ОДН на ГВС",
		"ОДН на электро",
		"ОДН на водоотв",
		"Содержание",
		"Электроэнергия",
	}

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

	// -- Для проверки печатаем тарифы на экран
	for idx, el := range tariffs {
		fmt.Println(idx, "\t", el)
	}
	fmt.Println()

	// Перебираем входной документ построчно
	rows, err := lastochka_file.GetRows("Жильцы")
	if err != nil {
		fmt.Println(err)
		return
	}

	type Flat struct {
		number int
		owner  string
		area   float64
		power  int
	}

	var CurrentFlat Flat
	var House []Flat
	var power, current_month, prev_month int

	// Перебираем строки, заносим в структуру и её в структуру общую
	for idx, rows := range rows {
		if idx < 129 {
			current_month, _ = strconv.Atoi(rows[month_cell])
			prev_month, _ = strconv.Atoi(rows[month_cell - 1])
			power = current_month - prev_month
			CurrentFlat.number, _ = strconv.Atoi(rows[0])
			CurrentFlat.owner = rows[1] + " " + rows[2] + " " + rows[3]
			CurrentFlat.area, _ = strconv.ParseFloat(rows[4], 64)
			CurrentFlat.power = power
			fmt.Println(CurrentFlat)
			House = append(House, CurrentFlat)
		} else {
			break
		}
	}
	fmt.Println("-=+++++=-")
	fmt.Println()
	fmt.Println(House[33])

	// szs_file, err := excelize.OpenFile("SZS1.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	// rec_file, err := excelize.OpenFile("SZS_rec.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	//  Создаем массив с ячейками в нашей таблице, в которой хранятся тарифы

	//  Создаем массив с названиями тарифов

	//  Получаем значения тарифов и формируем карту [Тариф: значение]

	//-----------------------------------------------------------------------------------------
	// // Формируем карту нужных полей в выходном документе
	// out_col := map[string]string{
	// 	"FACTP":   "L",
	// 	"TARIF":   "N",
	// 	"PRIZN":   "P",
	// 	"FACTOP":  "V",
	// 	"FACTOP2": "W",
	// }

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
