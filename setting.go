package main

import (
	"encoding/json"
	"fmt"
	"os"
)

type setting struct {
	In_file    string
	In_tariffs string
	In_owners  string
	T_odn_hvs  string
	T_odn_gvs  string
	T_odn_elec string
	T_odn_voda string
	T_soderzh  string
	T_electro  string
	Out_file   string
	Out_sheet  string
}

var cfg setting

func init() {
	// Открываем файл
	file, err := os.Open("setting.cfg")
	if err != nil {
		fmt.Println(err.Error())
		panic("ОШИБКА: Не удалось открыть файл конфигурации")
	}
	defer file.Close()

	stat, err := file.Stat()
	if err != nil {
		fmt.Println(err.Error())
		panic("ОШИБКА: Не удалось прочитать информацию о файле конфигурации")
	}

	readByte := make([]byte, stat.Size())

	_, err = file.Read(readByte)
	if err != nil {
		fmt.Println(err.Error())
		panic("ОШИБКА: Не удалось прочитать файл конфигурации")
	}

	err = json.Unmarshal(readByte, &cfg)
	if err != nil {
		fmt.Println(err.Error())
		panic("ОШИБКА: Не удалось считать данные файла конфигурации")
	}

}
