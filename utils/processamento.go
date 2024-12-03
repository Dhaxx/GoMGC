package utils

import (
	"bufio"
	"fmt"
	"log"
	"math/rand"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func regredirDias() []string {
	dataAtual := time.Now()
	var dias []string

	for i := 0; i <= 29; i++ {
		diaAnterior := dataAtual.Add(-24 * time.Hour)
		dataAtualFormatada := diaAnterior.Format("02/01/2006")
		dias = append(dias, dataAtualFormatada)
		dataAtual = diaAnterior
	}

	return dias
}

func PreencherPlanilha() {
	planilha, err := excelize.OpenFile("C:\\Programação\\GoMGC\\utils\\Glicemia.xlsx")
	if err != nil {
		log.Fatalf("Erro ao abrir a planilha: %v", err)
	}
	defer func() {
		filePath := filepath.Join("..", "Planilha-Preenchida.xlsx")
		// fmt.Print(filePath)

		err = planilha.SaveAs(filePath)
		if err != nil {
			log.Panicf("Erro ao salvar planilha modificada: %v", err)
		}

		if err = planilha.Close(); err != nil {
			log.Fatalf("Erro ao fechar planilha: %v", err)
		}

		cmd := exec.Command("C:\\Windows\\System32\\cmd.exe", "/c", "start", filePath)
		err = cmd.Start()
		if err != nil {
			log.Fatalf("Erro ao abrir a planilha preenchida: %v", err)
}
	}()

	const totalRows int = 34
	const totalCols int = 16

	var dia int
	var dias []string = regredirDias()
	planilhaAtual := "Planilha1"
	var trocouPlanilha bool

	for row := 6; row <= totalRows; row += 2 {
		for col := 1; col <= totalCols; col++ {
			// Converte índice x na letra da coluna correlata
			letraColuna, err := excelize.ColumnNumberToName(col)
			if err != nil {
				log.Fatalf("Erro ao converter índice para letra da coluna: %v", err)
			}
			stringLinha := strconv.Itoa(row)

			// Verifica se é coluna das datas
			if letraColuna == "A" {
				celula := letraColuna + stringLinha
				err = planilha.SetCellValue(planilhaAtual, celula, dias[dia])
				if err != nil {
					log.Fatalf("Erro ao tentar preencher célular: %v", err)
				}
				fmt.Printf("\n====%s====\n", dias[dia])
				dia += 1
			} else {
				var glicemia uint64

				if letraColuna == "C" {
					fmt.Print("Glicemia em Jejum: ")
				} else if letraColuna == "G" {
					fmt.Print("Glicemia pós Almoço: ")
				} else if letraColuna == "M" {
					fmt.Print("Glicemia pós Jantar: ")
				} else {
					continue
				}

				glicemia = receberGlicemia()
				doseInsulina(glicemia, letraColuna, stringLinha, planilha, planilhaAtual)

				celula := letraColuna + stringLinha
				err = planilha.SetCellValue(planilhaAtual, celula, glicemia)
				if err != nil {
					log.Fatalf("Erro ao salvar valor na célula: %v", err)
				}
			}
		}

		if row >= totalRows && !trocouPlanilha {
			if !trocouPlanilha {
				row = 4
			}
			planilhaAtual = "Planilha2"
			trocouPlanilha = true
			continue
		}
	}
}

func receberGlicemia() uint64 {
	reader := bufio.NewReader(os.Stdin)
	input, err := reader.ReadString('\n')
	if err != nil {
		log.Panicf("Erro ao ler a entrada: %v\n", err)
	}
	input = strings.TrimSpace(input)

	if input == "" {
		r := rand.New(rand.NewSource(time.Now().UnixNano()))

		// Intervalo desejado
		min := 50
		max := 300

		// Gerar número aleatório no intervalo [min, max]
		randomNumber := r.Intn(max-min+1) + min
		input = strconv.Itoa(randomNumber)
	}
	inputUint, err := strconv.ParseUint(input, 10, 64)
	if err != nil {
		log.Panicf("Erro ao converter entrada: %v", err)
	}

	return inputUint
}

func doseInsulina(glicemia uint64, letraColuna string, stringLinha string, planilha *excelize.File, planilhaAtual string) {
	var dosagem string
	var celula string

	if glicemia < 100 {
		dosagem = "-"
	} else if glicemia <= 120 {
		dosagem = "2uI"
	} else if glicemia <= 140 {
		dosagem = "4uI"
	} else if glicemia <= 160 {
		dosagem = "6uI"
	} else if glicemia <= 180 {
		dosagem = "8uI"
	} else if glicemia <= 200 {
		dosagem = "10uI"
	} else if glicemia <= 220 {
		dosagem = "8uI"
	} else {
		dosagem = "12uI"
	}

	if letraColuna == "C" {
		celula = "E" + stringLinha
	} else if letraColuna == "G" {
		celula = "H" + stringLinha
	} else if letraColuna == "M" {
		celula = "N" + stringLinha
	}

	err := planilha.SetCellValue(planilhaAtual, celula, dosagem)
	if err != nil {
		log.Panicf("Erro ao calcular dosagem de insulina: %v", err)
	}
}
