package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// Функция выбора месяца, возвращяет ячейку таблицы для дальнейшей обработки
func select_month_cell() int {

	var month int

	fmt.Println("Введите номер месяца, на который делаем расчет:")
	fmt.Println("________________")
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
	fmt.Println("================")
	fmt.Print("-> ")
	fmt.Scanln(&month)
	fmt.Println()

	for (month < 1) || (month > 12) {
		fmt.Println("ERROR: Неправильно введён месяц!")
		fmt.Print("-> ")
		fmt.Scanln(&month)
	}

	month_index := [13]int{15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27}

	return month_index[month]
}

// Функция генерации тарифов в 'map'
func generate_tariffs(file *excelize.File) map[string]float64 {
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

	tariff := make(map[string]float64)

	for idx, el := range tariff_cells {
		cell, err := file.GetCellValue("Квитанции_чистые", el)
		if err != nil {
			log.Fatal(err)
		}
		if cellfloat, err := strconv.ParseFloat(cell, 64); err == nil {
			tariff[tariff_names[idx]] = cellfloat
		} else {
			log.Fatal(err)
		}
	}
	return tariff
}

func main() {
	//  Открываем необходимые файлы (Ласточка, льготники и копию льготников(надо ли?))
	lastochka_file, err := excelize.OpenFile("Ласточка 2021.xlsm")
	if err != nil {
		log.Fatal(err)
	}

	month_cell := select_month_cell() // Получаем месяц, с которым работаем

	tariffs := generate_tariffs(lastochka_file) // Генерация 'map' с тарифами

	// -- TEST -- Для проверки печатаем тарифы на экран
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
	for idx, row := range rows {
		if idx < 129 {
			current_month, _ = strconv.Atoi(row[month_cell])
			prev_month, _ = strconv.Atoi(row[month_cell-1])
			power = current_month - prev_month
			CurrentFlat.number, _ = strconv.Atoi(row[0])
			CurrentFlat.owner = row[1] + " " + row[2] + " " + row[3]
			CurrentFlat.area, _ = strconv.ParseFloat(row[4], 64)
			CurrentFlat.power = power
			//fmt.Println(CurrentFlat)
			House = append(House, CurrentFlat)
			if idx == 20 {
				//fmt.Printf("Последнее показание RAW: %v\t Предыдущее показание RAW: %v\n", row[15], row[14])
				fmt.Printf("Последнее показание: %v\t Предыдущее показание: %v\n", current_month, prev_month)
				fmt.Printf("Кв: %v\t Владелец: %v\t Площадь: %v\t кВт: %v\n", House[20].number, House[20].owner, House[20].area, House[20].power)
			}
		} else {
			break
		}
	}
	fmt.Println("-=+++++=-")
	fmt.Println()
	fmt.Printf("Кв: %v\t тек. показ: -- пред. показ: -- кол-во кВт: %v\n", House[20].number, House[20].power)

	// // szs_file, err := excelize.OpenFile("SZS1.xlsx")
	// // if err != nil {
	// // 	log.Fatal(err)
	// // }

	rec_file, err := excelize.OpenFile("SZS_rec.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	//  Создаем массив с ячейками в нашей таблице, в которой хранятся тарифы

	//  Создаем массив с названиями тарифов

	//  Получаем значения тарифов и формируем карту [Тариф: значение]

	//-----------------------------------------------------------------------------------------
	// // Формируем карту нужных полей в выходном документе
	out_col := map[string]string{
		"FACTP":   "L",
		"TARIF":   "N",
		"PRIZN":   "P",
		"FACTOP":  "V",
		"FACTOP2": "W",
	}

	//
	out_rows, err := rec_file.GetRows("Лист1")
	if err != nil {
		fmt.Println(err)
		return
	}
	// // Разбираем каждую строку и вносим значения тарифов в выходную таблицу
	var kv int

	for idx, row := range out_rows {
		cell_Tarif := (out_col["TARIF"] + strconv.Itoa(idx+1))
		cell_Factp := (out_col["FACTP"] + strconv.Itoa(idx+1))
		cell_Factop := (out_col["FACTOP"] + strconv.Itoa(idx+1))
		cell_Factop2 := (out_col["FACTOP2"] + strconv.Itoa(idx+1))
		cell_Prizn := (out_col["PRIZN"] + strconv.Itoa(idx+1))
		// Ячейка с индексом 5 - номер квартиры в выходном документе
		kv, _ = strconv.Atoi(row[5])
		fmt.Println(row[5])
		if row[7] == "ОДН на ХВС" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на ХВС"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["ОДН на ХВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["ОДН на ХВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["ОДН на ХВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "ОДН на ГВС" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на ГВС"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["ОДН на ГВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["ОДН на ГВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["ОДН на ГВС"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "ОДН на водоотведение" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на водоотв"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["ОДН на водоотв"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["ОДН на водоотв"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["ОДН на водоотв"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Электрическая энергия на общедомовые нужды" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на электро"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["ОДН на электро"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["ОДН на электро"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["ОДН на электро"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Содержание жилья" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["Содержание"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["Содержание"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["Содержание"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["Содержание"] * House[kv].area))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Э: МЖД с ЦГВС и электроплитами" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["Электроэнергия"])
			rec_file.SetCellValue("Лист1", cell_Factp, (tariffs["Электроэнергия"] * float64(House[kv].power)))
			rec_file.SetCellValue("Лист1", cell_Factop, (tariffs["Электроэнергия"] * float64(House[kv].power)))
			rec_file.SetCellValue("Лист1", cell_Factop2, (tariffs["Электроэнергия"] * float64(House[kv].power)))
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else {
			if idx > 0 {
				rec_file.SetCellValue("Лист1", cell_Prizn, 1)
			}
		}
	}
	//
	// //  Сохраняем выходной файл
	if err := rec_file.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

}
