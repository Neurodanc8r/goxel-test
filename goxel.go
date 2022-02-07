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
		cfg.T_odn_hvs,
		cfg.T_odn_gvs,
		cfg.T_odn_elec,
		cfg.T_odn_voda,
		cfg.T_soderzh,
		cfg.T_electro,
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
		cell, err := file.GetCellValue(cfg.In_tariffs, el)
		if err != nil {
			log.Fatal(err)
		}
		if cellfloat, err := strconv.ParseFloat(cell, 64); err == nil {
			tariff[tariff_names[idx]] = cellfloat
		} else {
			log.Fatal(err)
		}
	}
	fmt.Printf("INFO: Чтение и формирование тарифов \t- ОК\n")
	return tariff
}

type Flat struct {
	number int
	owner  string
	area   float64
	power  int
}

// Функция читает инфо из входного файла и формирует срез структур 'House'
func read_gen_flat_info(file *excelize.File, month_cell int) []Flat {
	// Перебираем входной документ построчно
	rows, err := file.GetRows(cfg.In_owners)
	if err != nil {
		log.Fatal(err)
	}

	var CurrentFlat Flat
	var House []Flat
	var power, current_month, prev_month int

	// Перебираем строки, заносим в структуру и её в структуру общую
	for idx, row := range rows {
		if idx < 8 || idx == 88 || idx == 89 || idx == 107 || idx == 124 {
			CurrentFlat.number = 0
			CurrentFlat.owner = "void"
			CurrentFlat.area = 0
			CurrentFlat.power = 0
			House = append(House, CurrentFlat)
			continue
		}
		if idx < 129 {
			if month_cell > len(row) {
				fmt.Println("ОШИБКА: индекс месяца превышает длину строки!")
				log.Fatal(month_cell)
			}
			current_month, err = strconv.Atoi(row[month_cell])
			if err != nil {
				if row[month_cell] == "" {
					current_month = 0
				} else {
					fmt.Println("ОШИБКА: текущий месяц", idx, err)
				}
			}
			prev_month, err = strconv.Atoi(row[month_cell-1])
			if err != nil {
				if row[month_cell-1] == "" {
					prev_month = 0
				} else {
					fmt.Println("ОШИБКА: пред. месяц", idx, err)
				}
			}
			power = current_month - prev_month
			if power < 0 {
				fmt.Println("ОШИБКА: Отрицательное значение расхода э/энергии !!!")
			}
			CurrentFlat.number, err = strconv.Atoi(row[0])
			if err != nil {
				fmt.Println("Ошибка - номер квартиры")
				log.Fatal(err)
			}
			CurrentFlat.owner = row[1] + " " + row[2] + " " + row[3]
			CurrentFlat.area, err = strconv.ParseFloat(row[4], 64)
			if err != nil {
				log.Fatal(err)
			}
			CurrentFlat.power = power
			House = append(House, CurrentFlat)

		} else {
			break
		}
	}
	fmt.Printf("INFO: Чтение информации по квартирам \t- ОК\n")
	return House
}

// Функция записи данных в выходной файл
func record_out(rec_file *excelize.File, House []Flat, tariffs map[string]float64) {
	// // Формируем карту нужных полей в выходном документе
	out_col := map[string]string{
		"FACTP":   "L",
		"TARIF":   "N",
		"PRIZN":   "P",
		"FACTOP":  "V",
		"FACTOP2": "W",
	}

	out_rows, err := rec_file.GetRows(cfg.Out_sheet)
	if err != nil {
		fmt.Println("ОШИБКА: Ошибка чтения выходного файла !")
		log.Fatal(err)
	}
	// // Разбираем каждую строку и вносим значения тарифов в выходную таблицу
	var kv int
	var odn_hvs, odn_gvs, odn_voda, odn_electro, soderzh, electro float64

	for idx, row := range out_rows {
		cell_Tarif := (out_col["TARIF"] + strconv.Itoa(idx+1))
		cell_Factp := (out_col["FACTP"] + strconv.Itoa(idx+1))
		cell_Factop := (out_col["FACTOP"] + strconv.Itoa(idx+1))
		cell_Factop2 := (out_col["FACTOP2"] + strconv.Itoa(idx+1))
		cell_Prizn := (out_col["PRIZN"] + strconv.Itoa(idx+1))
		// Ячейка с индексом 5 - номер квартиры в выходном документе
		if row[5] == "KV" {
			continue
		} else {
			kv, err = strconv.Atoi(row[5]) // Читаем номер квартиры
			if err != nil {
				fmt.Println("ОШИБКА: Ошибка чтения номера квартиры выходного файла !")
				log.Fatal(err)
			}
		}

		odn_hvs = tariffs["ОДН на ХВС"] * House[kv].area // Считаем ОДН на ХВС
		if odn_hvs <= 0 {
			fmt.Println("ОШИБКА: отрицательный или нулевой тариф ОДН на ХВС !")
		}
		odn_gvs = tariffs["ОДН на ГВС"] * House[kv].area // Считаем ОДН на ГВС
		if odn_gvs <= 0 {
			fmt.Println("ОШИБКА: отрицательный или нулевой тариф ОДН на ГВС !")
		}
		odn_voda = tariffs["ОДН на водоотв"] * House[kv].area // Считаем ОДН на водоотведение
		if odn_voda <= 0 {
			fmt.Println("ОШИБКА: отрицательный или нулевой тариф ОДН на водоотведение !")
		}
		odn_electro = tariffs["ОДН на электро"] * House[kv].area
		if odn_electro <= 0 {
			fmt.Println("ОШИБКА: отрицательный или нулевой тариф ОДН на э/энергию !")
		}
		soderzh, err = strconv.ParseFloat(fmt.Sprintf("%.2f", tariffs["Содержание"]*House[kv].area), 64)
		if soderzh <= 0 {
			fmt.Println("ОШИБКА: отрицательное или нулевое значение суммы на содержание жилья !")
		}
		if err != nil {
			fmt.Println("ОШИБКА: Ошибка преобразования (округления) суммы на содержание жилья выходного файла !")
			log.Fatal(err)
		}
		electro, err = strconv.ParseFloat(fmt.Sprintf("%.2f", tariffs["Электроэнергия"]*float64(House[kv].power)), 64)
		if electro <= 0 {
			fmt.Println("ОШИБКА: отрицательное или нулевое значение суммы потребленной э/энергии !")
		}
		if err != nil {
			fmt.Println("ОШИБКА: Ошибка преобразования (округления) э/энергии выходного файла !")
			log.Fatal(err)
		}

		if row[7] == "ОДН на ХВС" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на ХВС"])
			rec_file.SetCellValue("Лист1", cell_Factp, odn_hvs)
			rec_file.SetCellValue("Лист1", cell_Factop, odn_hvs)
			rec_file.SetCellValue("Лист1", cell_Factop2, odn_hvs)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "ОДН на ГВС" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на ГВС"])
			rec_file.SetCellValue("Лист1", cell_Factp, odn_gvs)
			rec_file.SetCellValue("Лист1", cell_Factop, odn_gvs)
			rec_file.SetCellValue("Лист1", cell_Factop2, odn_gvs)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "ОДН на водоотведение" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на водоотв"])
			rec_file.SetCellValue("Лист1", cell_Factp, odn_voda)
			rec_file.SetCellValue("Лист1", cell_Factop, odn_voda)
			rec_file.SetCellValue("Лист1", cell_Factop2, odn_voda)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Электрическая энергия на общедомовые нужды" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["ОДН на электро"])
			rec_file.SetCellValue("Лист1", cell_Factp, odn_electro)
			rec_file.SetCellValue("Лист1", cell_Factop, odn_electro)
			rec_file.SetCellValue("Лист1", cell_Factop2, odn_electro)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Содержание жилья" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["Содержание"])
			rec_file.SetCellValue("Лист1", cell_Factp, soderzh)
			rec_file.SetCellValue("Лист1", cell_Factop, soderzh)
			rec_file.SetCellValue("Лист1", cell_Factop2, soderzh)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else if row[7] == "Э: МЖД с ЦГВС и электроплитами" {
			rec_file.SetCellValue("Лист1", cell_Tarif, tariffs["Электроэнергия"])
			rec_file.SetCellValue("Лист1", cell_Factp, electro)
			rec_file.SetCellValue("Лист1", cell_Factop, electro)
			rec_file.SetCellValue("Лист1", cell_Factop2, electro)
			rec_file.SetCellValue("Лист1", cell_Prizn, 1)
		} else {
			if idx > 0 {
				rec_file.SetCellValue("Лист1", cell_Prizn, 1)
			}
		}
	}
	fmt.Printf("INFO: Запись в выходной файл \t\t- ОК\n")
}

func main() {
	// Открываем входной файл с информацией (Ласточка) {новый будет Ласточка 2022_v1.xlsm}
	lastochka_file, err := excelize.OpenFile(cfg.In_file)
	if err != nil {
		fmt.Println("ОШИБКА: Ошибка чтения входного файла ! (Ласточка)")
		log.Fatal(err)
	}

	month_cell := select_month_cell() // Получаем месяц, с которым работаем

	tariffs := generate_tariffs(lastochka_file) // Генерация 'map' с тарифами

	House := read_gen_flat_info(lastochka_file, month_cell) // Получаем и храним в памяти информацию по квартирам

	// Открываем выходной файл по льготникам
	rec_file, err := excelize.OpenFile(cfg.Out_file)
	if err != nil {
		fmt.Println("ОШИБКА: Ошибка чтения выходного файла ! (по льготникам)")
		log.Fatal(err)
	}

	record_out(rec_file, House, tariffs) // Пишем в выходной файл (льготники)

	// Сохраняем выходной файл
	if err := rec_file.SaveAs(cfg.Out_file); err != nil {
		fmt.Println("ОШИБКА: Ошибка записи выходного файла ! (по льготникам)")
		fmt.Println(err)
	} else {
		fmt.Printf("INFO: Сохранение выходного файла \t- OK\n\n")
	}

	// -- TEST -- Для проверки печатаем тарифы на экран
	for idx, el := range tariffs {
		fmt.Println(idx, "\t", el)
	}
	fmt.Println()
	fmt.Println("Завершено успешно !")
	fmt.Println("В поля фактической оплаты внесены значения оплаты начисленной. Учет задолженностей не ведется.")

}
