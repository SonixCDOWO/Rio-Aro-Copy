package main

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"html/template"
	"io"
	"io/ioutil"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	wkhtml "github.com/SebastiaanKlippert/go-wkhtmltopdf" // Para PDF
)

// Google
// Roles
const (
	RoleAdmin    = "admin"
	RoleOperator = "operator"
)

// Configuración de usuarios permitidos (Hardcoded como pediste)
var authorizedUsers = map[string]string{
	"sebastianalbertoruizvargas@gmail.com": RoleAdmin,
	"silvaguerrajosecarlos@gmail.com":      RoleAdmin,
	"operador1@gmail.com":                  RoleOperator,
}

type UserSession struct {
	Email   string        `json:"email"`
	Name    string        `json:"name"` // <-- Asegúrate de que esta línea existe
	Picture string        `json:"picture"`
	Role    string        `json:"role"`
	Token   *oauth2.Token // Para mantener la conexión con Drive
}

var (
	googleOauthConfig *oauth2.Config
	// Simulación de sesión simple (en producción usa cookies seguras o JWT)
	activeSession *UserSession
)

// DriveService envuelve la conexión con Google
type DriveSync struct {
	Service *drive.Service
	FileID  string // El ID del archivo Excel en la nube
}

func (ds *DriveSync) GetRemoteMetadata() (*drive.File, error) {
	return ds.Service.Files.Get(ds.FileID).Fields("modifiedTime", "name", "id").Do()
}

// Función para subir el archivo (Commit/Push)
func (ds *DriveSync) PushLocalToRemote() error {
	f, err := os.Open(EXCEL_FILE)
	if err != nil {
		return err
	}
	defer f.Close()

	driveFile := &drive.File{}
	_, err = ds.Service.Files.Update(ds.FileID, driveFile).Media(f).Do()
	return err
}

// Función para descargar el archivo (Pull)
func (ds *DriveSync) PullRemoteToLocal() error {
	resp, err := ds.Service.Files.Get(ds.FileID).Download()
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	out, err := os.Create(EXCEL_FILE)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, resp.Body)
	return err
}
func AuthMiddleware(next http.HandlerFunc, requiredRole string) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		if activeSession == nil {
			http.Redirect(w, r, "/login", http.StatusTemporaryRedirect)
			return
		}

		if requiredRole == RoleAdmin && activeSession.Role != RoleAdmin {
			http.Error(w, "Acceso denegado: Se requieren permisos de Administrador", http.StatusForbidden)
			return
		}

		next.ServeHTTP(w, r)
	}
}
func init() {
	// Carga las credenciales desde el archivo JSON
	data, err := ioutil.ReadFile("credentials.json")
	if err != nil {
		fmt.Printf("Error leyendo credentials.json: %v\n", err)
		return
	}

	conf, err := google.ConfigFromJSON(data,
		drive.DriveFileScope,
		"https://www.googleapis.com/auth/userinfo.email",
		"https://www.googleapis.com/auth/userinfo.profile",
	)
	if err != nil {
		fmt.Printf("Error configurando Google OAuth: %v\n", err)
		return
	}
	googleOauthConfig = conf
}

// Excel
const EXCEL_FILE = "CENSO GENERAL NUEVO.xlsx"
const PRIMERA_HOJA = "CENSO"

var historyLogs []string = []string{
	"Sistema iniciado: " + time.Now().Format("2006-01-02 15:04:05"),
}

// --- INICIO DE LA MODIFICACIÓN ---
// Define qué columnas usarán una búsqueda "Contains" (contiene).
// Las columnas que NO estén en este mapa usarán una búsqueda exacta (==).
// Asegúrate de que los nombres aquí coincidan EXACTAMENTE con las cabeceras de tu Excel (después de limpiar espacios).
var containsSearchColumns = map[string]bool{
	"Nombre completo":     true,
	"Cedula de identidad": true,
	// Si tienes otras columnas de texto libre (ej: "Dirección", "Observaciones"), añádelas aquí.
	// "Dirección": true,
}

// --- FIN DE LA MODIFICACIÓN ---

// Galeria
type GalleryData struct {
	Images   []string
	Messages []string
}

// Estructura para un nodo en la vista de árbol (jerarquía)
type TreeNode struct {
	Text     string      `json:"text"`     // El texto que se mostrará (ej: "MZA", "15", "A1")
	Type     string      `json:"type"`     // <-- 1. AÑADIMOS ESTE CAMPO
	Children []*TreeNode `json:"children"` // Nodos hijos (recursivo)
	State    struct {
		Opened bool `json:"opened"`
	} `json:"state"` // Para que los nodos aparezcan cerrados por defecto
}

// Estructura para una persona individual
type Person struct {
	Parentesco string `json:"parentesco"`
	Nombres    string `json:"nombres"`
	Documento  string `json:"documento"`
}

const uploadDir = "assets/imagenes"

// Abre la URL en el navegador predeterminado del sistema operativo.
func openBrowser(url string) {
	var err error

	switch runtime.GOOS {
	case "windows":
		err = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
	case "darwin":
		err = exec.Command("open", url).Start()
	case "linux":
		err = exec.Command("xdg-open", url).Start()
	default:
		fmt.Println("No se pudo abrir el navegador automáticamente.")
	}

	if err != nil {
		fmt.Println("Error al abrir el navegador:", err)
	}
}

// ------------------- MANEJO DEL EXCEL -------------------------
// deleteRowHandler elimina una fila específica del archivo Excel.
func deleteRowHandler(w http.ResponseWriter, r *http.Request) {
	fmt.Println("--- LOG: Endpoint /api/delete-row invocado. ---")

	var req struct {
		// El frontend enviará el número de fila en el JSON
		Row int `json:"__row"`
	}

	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		fmt.Printf("--- ERROR: No se pudo decodificar el payload JSON: %v ---\n", err)
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	fmt.Printf("--- LOG: Solicitud para eliminar la fila número: %d ---\n", req.Row)

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		fmt.Printf("--- ERROR: No se pudo abrir el archivo Excel: %v ---\n", err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	// Usamos la función de la librería excelize para remover la fila
	if err := f.RemoveRow(PRIMERA_HOJA, req.Row); err != nil {
		fmt.Printf("--- ERROR: No se pudo remover la fila %d del sheet: %v ---\n", req.Row, err)
		http.Error(w, "Error al remover la fila", http.StatusInternalServerError)
		return
	}

	fmt.Printf("--- LOG: Fila %d removida del sheet en memoria. Guardando archivo... ---\n", req.Row)

	if err := f.Save(); err != nil {
		fmt.Printf("--- ERROR: No se pudo guardar el archivo Excel tras la eliminación: %v ---\n", err)
		http.Error(w, "No se guardó el Excel", http.StatusInternalServerError)
		return
	}

	fmt.Println("--- LOG: ¡Archivo Excel guardado exitosamente! ---")
	w.WriteHeader(http.StatusOK)
}

func normalizeHeader(header string) string {
	// 1. Convertir a minúsculas
	lower := strings.ToLower(header)
	// 2. Compilar una expresión regular para encontrar cualquier cosa que NO sea una letra o un número
	reg := regexp.MustCompile("[^a-z0-9]+")
	// 3. Reemplazar esos caracteres con una cadena vacía
	return reg.ReplaceAllString(lower, "")
}

// bulkImportHandler (versión mejorada con mapeo inteligente)
func bulkImportHandler(w http.ResponseWriter, r *http.Request) {
	fmt.Println("--- LOG: Endpoint /api/bulk-import invocado. ---")

	var req struct {
		Datos []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		fmt.Printf("--- ERROR: No se pudo decodificar el payload JSON: %v ---\n", err)
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	fmt.Printf("--- LOG: Recibidas %d personas para importar.\n", len(req.Datos))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]
	nextRow := len(rows) + 1

	// --- LÓGICA DE MAPEO INTELIGENTE ---
	// Crear un mapa donde la clave es el nombre de la cabecera NORMALIZADO
	// y el valor es el índice de la columna original.
	normalizedHeaderMap := make(map[string]int)
	for i, h := range headers {
		normalized := normalizeHeader(h)
		normalizedHeaderMap[normalized] = i
		fmt.Printf("--- LOG (Mapeo Censo): '%s' -> '%s' en índice %d\n", h, normalized, i)
	}

	fmt.Println("--- LOG: Iniciando proceso de escritura en el Excel... ---")
	for _, persona := range req.Datos {
		for keyFromImport, val := range persona {
			// Normalizar la clave del archivo importado
			normalizedKeyFromImport := normalizeHeader(keyFromImport)

			// Buscar la columna correspondiente en nuestro mapa normalizado
			if colIndex, ok := normalizedHeaderMap[normalizedKeyFromImport]; ok {
				// Se encontró una coincidencia, escribir en la celda correcta
				cell, _ := excelize.CoordinatesToCellName(colIndex+1, nextRow)
				f.SetCellValue(PRIMERA_HOJA, cell, val)
			} else {
				fmt.Printf("--- AVISO: La columna importada '%s' (normalizada a '%s') no se encontró en el censo principal y será ignorada.\n", keyFromImport, normalizedKeyFromImport)
			}
		}
		nextRow++
	}

	if err := f.Save(); err != nil {
		fmt.Printf("--- ERROR: No se pudo guardar el archivo Excel: %v ---\n", err)
		http.Error(w, "No se guardó el Excel", http.StatusInternalServerError)
		return
	}

	fmt.Println("--- LOG: ¡Importación en bloque completada y archivo guardado! ---")
	w.WriteHeader(http.StatusOK)
}

// checkCedulasHandler (con logs detallados) recibe una lista de cédulas y devuelve las que ya existen.
func checkCedulasHandler(w http.ResponseWriter, r *http.Request) {
	fmt.Println("--- LOG (check-cedulas): Endpoint invocado. ---")
	var req struct {
		Cedulas []string `json:"cedulas"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}
	fmt.Printf("--- LOG (check-cedulas): Recibidas %d Cédulas para verificar.\n", len(req.Cedulas))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}
	rows, _ := f.GetRows(PRIMERA_HOJA)

	cedulaHeaderNormalized := normalizeHeader("Cedula de identidad")
	cedulaColIndex := -1
	for i, h := range rows[0] {
		if normalizeHeader(h) == cedulaHeaderNormalized {
			cedulaColIndex = i
			break
		}
	}
	if cedulaColIndex == -1 {
		fmt.Println("--- ERROR (check-cedulas): No se encontró la columna 'Cedula de identidad' en el archivo principal.")
		http.Error(w, "No se encontró la columna de Cedula", http.StatusInternalServerError)
		return
	}
	fmt.Printf("--- LOG (check-cedulas): Columna 'Cedula de identidad' encontrada en el índice %d.\n", cedulaColIndex)

	existingCedulas := make(map[string]bool)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if cedulaColIndex < len(row) {
			existingCedulas[row[cedulaColIndex]] = true
		}
	}
	fmt.Printf("--- LOG (check-cedulas): Se construyó un set con %d Cédulas existentes del archivo principal.\n", len(existingCedulas))

	var duplicates []string
	for _, cedula := range req.Cedulas {
		if _, exists := existingCedulas[cedula]; exists {
			duplicates = append(duplicates, cedula)
		}
	}

	fmt.Printf("--- LOG (check-cedulas): Verificación completada. Se encontraron %d duplicados.\n", len(duplicates))
	fmt.Printf("--- LOG (check-cedulas): Enviando respuesta al frontend: %v\n", duplicates)
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(duplicates)
}

// getPersonByCedulaHandler (con logs detallados) busca una persona por su cédula.
func getPersonByCedulaHandler(w http.ResponseWriter, r *http.Request) {
	cedulaToFind := r.URL.Query().Get("cedula")
	fmt.Printf("--- LOG (get-person): Endpoint invocado para buscar la cédula: %s ---\n", cedulaToFind)

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}
	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]

	cedulaHeaderNormalized := normalizeHeader("Cedula de identidad")
	cedulaColIndex := -1
	for i, h := range headers {
		if normalizeHeader(h) == cedulaHeaderNormalized {
			cedulaColIndex = i
			break
		}
	}
	if cedulaColIndex == -1 {
		fmt.Println("--- ERROR (get-person): No se encontró la columna 'Cedula de identidad'.")
		http.Error(w, "No se encontró la columna de Cedula", http.StatusInternalServerError)
		return
	}

	var personData map[string]string
	fmt.Printf("--- LOG (get-person): Buscando en el archivo principal...\n")
	for i, row := range rows[1:] {
		if cedulaColIndex < len(row) && row[cedulaColIndex] == cedulaToFind {
			fmt.Printf("--- LOG (get-person): ¡Coincidencia encontrada en la fila %d del Excel!\n", i+2)
			personData = make(map[string]string)
			for j, header := range headers {
				if j < len(row) {
					personData[header] = row[j]
				}
			}
			break
		}
	}

	if personData == nil {
		fmt.Println("--- LOG (get-person): Persona no encontrada. Enviando respuesta 404 Not Found.")
		http.NotFound(w, r)
		return
	}

	fmt.Println("--- LOG (get-person): Enviando datos de la persona encontrada al frontend.")
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(personData)
}

// Estructuras para las respuestas JSON
type ExcelResponse struct {
	Headers []string            `json:"headers"`
	Data    []map[string]string `json:"data"`
}

// Estructura para la respuesta del DataTables
type DTResponse struct {
	Draw            int                 `json:"draw"`
	RecordsTotal    int                 `json:"recordsTotal"`
	RecordsFiltered int                 `json:"recordsFiltered"`
	Data            []map[string]string `json:"data"`
}

// Obtiene las columnas del Excel y las devuelve como JSON
func getColumns(w http.ResponseWriter, r *http.Request) {
	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	// Leer solo la primera fila del sheet PRIMERA_HOJA
	row, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(row) == 0 {
		http.Error(w, "sheet vacío o no existe", 500)
		return
	}

	headers := row[0]

	// Limpiar espacios de las cabeceras antes de enviarlas al frontend
	cleanHeaders := make([]string, len(headers))
	for i, h := range headers {
		cleanHeaders[i] = strings.TrimSpace(h)
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(cleanHeaders) // Enviar las cabeceras limpias
}

// Leer datos del Excel y paginarlos para DataTables (FUNCIÓN CORREGIDA)
func getData(w http.ResponseWriter, r *http.Request) {
	// Filtro global
	search := strings.ToLower(r.URL.Query().Get("search[value]"))

	// Paginación
	start, _ := strconv.Atoi(r.URL.Query().Get("start"))
	length, _ := strconv.Atoi(r.URL.Query().Get("length"))
	draw, _ := strconv.Atoi(r.URL.Query().Get("draw"))

	// Leer los parámetros de filtro de columna
	filterColumnName := r.URL.Query().Get("filterColumn") // Ya viene limpio del frontend
	filterValue := strings.ToLower(r.URL.Query().Get("filterValue"))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	// Leer todas las filas de la hoja
	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", 500)
		return
	}

	// Leer cabecera para los keys
	keys := rows[0]

	// Encontrar el índice de la columna que queremos filtrar
	filterColumnIndex := -1
	if filterColumnName != "" {
		for i, key := range keys {
			// Comparar la cabecera limpia (TrimSpace) con el nombre que viene del frontend
			if strings.TrimSpace(key) == filterColumnName {
				filterColumnIndex = i
				break
			}
		}
	}

	type IndexedRow struct {
		Index int      // índice real en el Excel
		Cells []string // contenido de la fila
	}

	filtered := make([]IndexedRow, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		// 1. Comprobación del filtro global (search[value])
		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), search) {
					globalMatch = true
					break
				}
			}
		}

		// 2. Comprobación del filtro por columna (filterColumn / filterValue)
		columnMatch := false
		if filterValue == "" || filterColumnIndex == -1 {
			columnMatch = true // Si no hay filtro de columna, todas las filas coinciden
		} else if filterColumnIndex < len(row) {

			// --- INICIO DE LA CORRECCIÓN ---
			cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))

			// Decidir si usar "Contains" o "Exact Match"
			// 'filterColumnName' es el nombre limpio de la columna (ej: "Nombre completo")
			if containsSearchColumns[filterColumnName] {
				// Usar "Contains" para esta columna
				if strings.Contains(cellValue, filterValue) {
					columnMatch = true
				}
			} else {
				// Usar "Exact Match" para esta columna
				if cellValue == filterValue {
					columnMatch = true
				}
			}
			// --- FIN DE LA CORRECCIÓN ---
		}

		// 3. Decisión final: La fila debe coincidir con AMBOS filtros
		if globalMatch && columnMatch {
			filtered = append(filtered, IndexedRow{Index: i + 1, Cells: row}) // +1 porque Excel empieza en 1
		}
	}

	data := make([]map[string]string, 0, length)
	for i := start; i < len(filtered) && len(data) < length; i++ {
		row := filtered[i]
		rec := map[string]string{}
		rec["__row"] = strconv.Itoa(row.Index)

		for j, key := range keys {
			val := ""
			if j < len(row.Cells) {
				val = row.Cells[j]
			}
			rec[strings.TrimSpace(key)] = val // Usar la cabecera limpia como clave
		}
		data = append(data, rec)
	}

	resp := DTResponse{
		Draw:            draw,
		RecordsTotal:    len(rows) - 1,
		RecordsFiltered: len(filtered),
		Data:            data,
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(resp)
}

// Estructura para la solicitud de actualización del Excel
type UpdateRequest struct {
	Datos []map[string]string `json:"datos"`
}

// updateExcelData
func updateExcelData(w http.ResponseWriter, r *http.Request) {
	fmt.Println("--- LOG: Endpoint /api/update-excel invocado. ---")

	var req struct {
		Datos []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		fmt.Printf("--- ERROR: No se pudo decodificar el payload JSON: %v ---\n", err)
		http.Error(w, "payload inválido", http.StatusBadRequest)
		return
	}

	fmt.Printf("--- LOG: Payload recibido del frontend: %+v\n", req.Datos)

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		fmt.Printf("--- ERROR: No se pudo abrir el archivo Excel: %v ---\n", err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]
	nextAvailableRow := len(rows) + 1

	for _, fila := range req.Datos {
		rowNumStr := fila["__row"]
		rowNum, err := strconv.Atoi(rowNumStr)

		if err != nil {
			fmt.Printf("--- LOG: Detectada nueva persona. Se agregará en la fila %d\n", nextAvailableRow)
			for colIndex, key := range headers {
				cleanKey := strings.TrimSpace(key)
				if val, ok := fila[cleanKey]; ok {
					cell, _ := excelize.CoordinatesToCellName(colIndex+1, nextAvailableRow)
					fmt.Printf("    -> Escribiendo en celda %s: '%s'\n", cell, val)
					f.SetCellValue(PRIMERA_HOJA, cell, val)
				}
			}
			nextAvailableRow++
		} else {
			fmt.Printf("--- LOG: Detectada persona existente. Se actualizará la fila %d\n", rowNum)
			for colIndex, key := range headers {
				cleanKey := strings.TrimSpace(key)
				if val, ok := fila[cleanKey]; ok {
					cell := fmt.Sprintf("%s%d", columnLetter(colIndex), rowNum)
					fmt.Printf("    -> Escribiendo en celda %s: '%s'\n", cell, val)
					f.SetCellValue(PRIMERA_HOJA, cell, val)
				}
			}
		}
	}

	fmt.Println("--- LOG: Intentando guardar los cambios en el archivo Excel... ---")
	if err := f.Save(); err != nil {
		fmt.Printf("--- ERROR: No se pudo guardar el archivo Excel: %v ---\n", err)
		http.Error(w, "no se guardó el Excel", http.StatusInternalServerError)
		return
	}

	fmt.Println("--- LOG: ¡Archivo Excel guardado exitosamente! ---")
	w.WriteHeader(http.StatusOK)
}

// exportToExcel (FUNCIÓN CORREGIDA)
func exportToExcel(w http.ResponseWriter, r *http.Request) {
	search := strings.ToLower(r.URL.Query().Get("search[value]"))
	filterColumnName := r.URL.Query().Get("filterColumn")
	filterValue := strings.ToLower(r.URL.Query().Get("filterValue"))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", 500)
		return
	}

	headers := rows[0]

	filterColumnIndex := -1
	if filterColumnName != "" {
		for i, key := range headers {
			if strings.TrimSpace(key) == filterColumnName {
				filterColumnIndex = i
				break
			}
		}
	}

	filteredRows := make([][]string, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), search) {
					globalMatch = true
					break
				}
			}
		}

		columnMatch := false
		if filterValue == "" || filterColumnIndex == -1 {
			columnMatch = true
		} else if filterColumnIndex < len(row) {
			// --- INICIO DE LA CORRECCIÓN ---
			cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))

			// Decidir si usar "Contains" o "Exact Match"
			if containsSearchColumns[filterColumnName] {
				// Usar "Contains"
				if strings.Contains(cellValue, filterValue) {
					columnMatch = true
				}
			} else {
				// Usar "Exact Match"
				if cellValue == filterValue {
					columnMatch = true
				}
			}
			// --- FIN DE LA CORRECCIÓN ---
		}

		if globalMatch && columnMatch {
			filteredRows = append(filteredRows, row)
		}
	}

	exportFile := excelize.NewFile()
	sheetName := "Reporte"
	index := exportFile.NewSheet(sheetName)
	exportFile.SetActiveSheet(index)

	for colIndex, header := range headers {
		cell := fmt.Sprintf("%s%d", columnLetter(colIndex), 1)
		exportFile.SetCellValue(sheetName, cell, header) // Exportar con la cabecera original
	}

	for rowIndex, rowData := range filteredRows {
		for colIndex, cellValue := range rowData {
			cell := fmt.Sprintf("%s%d", columnLetter(colIndex), rowIndex+2)
			exportFile.SetCellValue(sheetName, cell, cellValue)
		}
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename=reporte_habitantes.xlsx")

	if err := exportFile.Write(w); err != nil {
		http.Error(w, "no se pudo escribir el archivo Excel", http.StatusInternalServerError)
	}
}

// exportToPDF (FUNCIÓN CORREGIDA)
func exportToPDF(w http.ResponseWriter, r *http.Request) {
	search := strings.ToLower(r.URL.Query().Get("search[value]"))
	filterColumnName := r.URL.Query().Get("filterColumn")
	filterValue := strings.ToLower(r.URL.Query().Get("filterValue"))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", 500)
		return
	}

	allHeaders := rows[0]

	filterColumnIndex := -1
	if filterColumnName != "" {
		for i, key := range allHeaders {
			if strings.TrimSpace(key) == filterColumnName {
				filterColumnIndex = i
				break
			}
		}
	}

	selectedColumns := []string{"Nombre completo", "Cedula de identidad", "Edad", "Genero"}

	var displayHeaders []string
	headerIndexMap := make(map[string]int)
	for i, header := range allHeaders {
		cleanHeader := strings.TrimSpace(header)
		for _, selectedCol := range selectedColumns {
			if cleanHeader == selectedCol { // Comparar limpio
				displayHeaders = append(displayHeaders, cleanHeader)
				headerIndexMap[cleanHeader] = i
				break
			}
		}
	}

	filteredData := make([]map[string]string, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), search) {
					globalMatch = true
					break
				}
			}
		}

		columnMatch := false
		if filterValue == "" || filterColumnIndex == -1 {
			columnMatch = true
		} else if filterColumnIndex < len(row) {
			// --- INICIO DE LA CORRECCIÓN ---
			cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))

			// Decidir si usar "Contains" o "Exact Match"
			if containsSearchColumns[filterColumnName] {
				// Usar "Contains"
				if strings.Contains(cellValue, filterValue) {
					columnMatch = true
				}
			} else {
				// Usar "Exact Match"
				if cellValue == filterValue {
					columnMatch = true
				}
			}
			// --- FIN DE LA CORRECCIÓN ---
		}

		if globalMatch && columnMatch {
			rowData := make(map[string]string)
			for _, header := range displayHeaders {
				originalIndex := headerIndexMap[header]
				val := ""
				if originalIndex < len(row) {
					val = row[originalIndex]
				}
				rowData[header] = val
			}
			filteredData = append(filteredData, rowData)
		}
	}

	// --- Generación del HTML para el PDF ---
	data := struct {
		Headers      []string
		Rows         []map[string]string
		Search       string
		FilterColumn string
		FilterValue  string
		RowCount     int
	}{
		Headers:      displayHeaders,
		Rows:         filteredData,
		Search:       r.URL.Query().Get("search[value]"),
		FilterColumn: filterColumnName,
		FilterValue:  r.URL.Query().Get("filterValue"),
		RowCount:     len(filteredData),
	}

	htmlTemplate := `
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<title>Reporte de Habitantes</title>
		<style>
			body { font-family: Arial, sans-serif; margin: 20px; }
			h1 { text-align: center; color: #333; }
			h3 { text-align: center; color: #333; }
			p { text-align: center; color: #666; }
			table { width: 100%; border-collapse: collapse; margin-top: 20px; }
			th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
			th { background-color: #f2f2f2; }
			tr:nth-child(even) { background-color: #f9f9f9; }
		</style>
	</head>
	<body>
		<h1>Reporte de Habitantes de Río Aro</h1>
		{{if .Search}}
		<p>Filtrado global por: "{{.Search}}"</p>
		{{end}}
		{{if .FilterValue}}
		<p>Filtrado de columna "{{.FilterColumn}}" por: "{{.FilterValue}}"</p>
		{{end}}
		<h3>Cantidad de filas filtradas: {{.RowCount}}</h3>
		<table>
			<thead>
				<tr>
					{{range .Headers}}
					<th>{{.}}</th>
					{{end}}
				</tr>
			</thead>
			<tbody>
			{{range $rowIndex, $row := .Rows}}
				<tr>
				{{range $colIndex, $header := $.Headers}}
					<td>{{index $row $header}}</td>
				{{end}}
				</tr>
			{{end}}
			</tbody>

		</table>
	</body>
	</html>
	`

	tmpl, err := template.New("pdfReport").Parse(htmlTemplate)
	if err != nil {
		http.Error(w, "error al parsear el template HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	var htmlBuffer bytes.Buffer
	if err := tmpl.Execute(&htmlBuffer, data); err != nil {
		http.Error(w, "error al ejecutar el template HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// --- Conversión de HTML a PDF con wkhtmltopdf ---
	pdfg, err := wkhtml.NewPDFGenerator()
	if err != nil {
		http.Error(w, "no se pudo crear el generador de PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	pdfg.AddPage(wkhtml.NewPageReader(bytes.NewReader(htmlBuffer.Bytes())))
	pdfg.PageSize.Set(wkhtml.PageSizeA4)
	pdfg.Orientation.Set(wkhtml.OrientationPortrait)

	err = pdfg.Create()
	if err != nil {
		http.Error(w, "no se pudo generar el PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", "attachment; filename=reporte_habitantes.pdf")
	w.Header().Set("Content-Length", strconv.Itoa(len(pdfg.Bytes())))

	if _, err := w.Write(pdfg.Bytes()); err != nil {
		http.Error(w, "no se pudo escribir el archivo PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}
}

// convierte 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB, etc.
func columnLetter(idx int) string {
	var col string
	for idx >= 0 {
		col = string(rune('A'+(idx%26))) + col
		idx = idx/26 - 1
	}
	return col
}

// ------------------- INICIO DEL SERVIDOR -------------------------
// Estructura para las actividades del calendario
type Activity struct {
	ID          int    `json:"id"`
	Title       string `json:"title"`
	Description string `json:"description"`
	StartDate   string `json:"start_date"`
	EndDate     string `json:"end_date"`
	Time        string `json:"time"`
	Location    string `json:"location"`
	Image       string `json:"image"`
}

var activities = []Activity{
	{ID: 1, Title: "Reunión Consejo Comunal", StartDate: "2025-09-15", Time: "10:00", Location: "Salón Comunal", Image: ""},
	{ID: 2, Title: "Jornada de Vacunación", StartDate: "2025-09-20", EndDate: "2025-09-21", Description: "Jornada de vacunación para niños y adultos mayores.", Location: "Centro de Salud", Image: ""},
}
var lastActivityID = 2

func getActivitiesHandler(w http.ResponseWriter, r *http.Request) {
	events := make([]map[string]interface{}, 0, len(activities))
	for _, a := range activities {
		event := map[string]interface{}{
			"id":    a.ID,
			"title": a.Title,
			"start": a.StartDate,
			"end":   a.EndDate,
			"extendedProps": map[string]string{
				"description": a.Description,
				"time":        a.Time,
				"location":    a.Location,
				"image":       a.Image,
			},
		}
		if a.EndDate != "" {
			end, _ := time.Parse("2006-01-02", a.EndDate)
			end = end.Add(24 * time.Hour)
			event["end"] = end.Format("2006-01-02")
		}
		events = append(events, event)
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(events)
}

func addActivityHandler(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "No se pudo parsear el formulario", http.StatusBadRequest)
		return
	}

	var newActivity Activity
	newActivity.Title = r.FormValue("title")
	newActivity.Description = r.FormValue("description")
	newActivity.StartDate = r.FormValue("start_date")
	newActivity.EndDate = r.FormValue("end_date")
	newActivity.Time = r.FormValue("time")
	newActivity.Location = r.FormValue("location")

	file, header, err := r.FormFile("image")
	if err == nil {
		defer file.Close()
		filename := fmt.Sprintf("%d-%s", time.Now().UnixNano(), filepath.Base(header.Filename))
		ext := strings.ToLower(filepath.Ext(filename))
		if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
			http.Error(w, "Formato de imagen no permitido", http.StatusBadRequest)
			return
		}

		dst, err := os.Create(filepath.Join(uploadDir, filename))
		if err != nil {
			http.Error(w, "Error al guardar la imagen", http.StatusInternalServerError)
			return
		}
		defer dst.Close()
		io.Copy(dst, file)
		newActivity.Image = filename
	} else if err != http.ErrMissingFile {
		http.Error(w, "Error al procesar el archivo subido", http.StatusInternalServerError)
		return
	}

	lastActivityID++
	newActivity.ID = lastActivityID
	activities = append(activities, newActivity)

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

func editActivityHandler(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/activities/edit/")
	id, err := strconv.Atoi(idStr)
	if err != nil {
		http.Error(w, "ID inválido", http.StatusBadRequest)
		return
	}

	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "No se pudo parsear el formulario", http.StatusBadRequest)
		return
	}

	var updatedActivity Activity
	updatedActivity.Title = r.FormValue("title")
	updatedActivity.Description = r.FormValue("description")
	updatedActivity.StartDate = r.FormValue("start_date")
	updatedActivity.EndDate = r.FormValue("end_date")
	updatedActivity.Time = r.FormValue("time")
	updatedActivity.Location = r.FormValue("location")

	for i, a := range activities {
		if a.ID == id {
			activities[i].Title = updatedActivity.Title
			activities[i].Description = updatedActivity.Description
			activities[i].StartDate = updatedActivity.StartDate
			activities[i].EndDate = updatedActivity.EndDate
			activities[i].Time = updatedActivity.Time
			activities[i].Location = updatedActivity.Location

			file, header, err := r.FormFile("image")
			if err == nil {
				defer file.Close()
				if activities[i].Image != "" {
					os.Remove(filepath.Join(uploadDir, activities[i].Image))
				}
				filename := fmt.Sprintf("%d-%s", time.Now().UnixNano(), filepath.Base(header.Filename))
				ext := strings.ToLower(filepath.Ext(filename))
				if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
					http.Error(w, "Formato de imagen no permitido", http.StatusBadRequest)
					return
				}

				dst, err := os.Create(filepath.Join(uploadDir, filename))
				if err != nil {
					http.Error(w, "Error al guardar la nueva imagen", http.StatusInternalServerError)
					return
				}
				defer dst.Close()
				io.Copy(dst, file)
				activities[i].Image = filename
			} else if err != http.ErrMissingFile {
				http.Error(w, "Error al procesar la nueva imagen", http.StatusInternalServerError)
				return
			}
			break
		}
	}
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

func deleteActivityHandler(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/activities/delete/")
	id, err := strconv.Atoi(idStr)
	if err != nil {
		http.Error(w, "ID inválido", http.StatusBadRequest)
		return
	}

	for i, a := range activities {
		if a.ID == id {
			if a.Image != "" {
				err := os.Remove(filepath.Join(uploadDir, a.Image))
				if err != nil {
					fmt.Printf("Advertencia: no se pudo eliminar la imagen %s: %v\n", a.Image, err)
				}
			}
			activities = append(activities[:i], activities[i+1:]...)
			break
		}
	}
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

// Handler NUEVO para el Historial (corrige el error Unexpected token)
func getHistoryHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(historyLogs)
}

func main() {
	os.MkdirAll(uploadDir, os.ModePerm)
	// Rutas de Auth
	http.HandleFunc("/login", handleGoogleLogin)
	http.HandleFunc("/callback", handleGoogleCallback)
	http.HandleFunc("/api/user/status", getUserStatus)
	http.HandleFunc("/api/sync/status", getSyncStatus)
	http.HandleFunc("/api/sync/pull", handlePull)
	http.HandleFunc("/api/sync/push", handlePush)

	// Cerrar sesión
	http.HandleFunc("/logout", func(w http.ResponseWriter, r *http.Request) {
		activeSession = nil
		http.Redirect(w, r, "/", http.StatusSeeOther)
	})
	//  Rutas api
	http.HandleFunc("/api/activities", getActivitiesHandler)
	http.HandleFunc("/api/activities/add", addActivityHandler)
	http.HandleFunc("/api/activities/edit/", editActivityHandler)
	http.HandleFunc("/api/activities/delete/", deleteActivityHandler)
	http.HandleFunc("/api/delete-row", deleteRowHandler)
	http.HandleFunc("/importar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/importar.html")
	})
	http.HandleFunc("/api/history", getHistoryHandler)
	http.HandleFunc("/api/bulk-import", bulkImportHandler)
	http.HandleFunc("/api/check-cedulas", checkCedulasHandler)
	http.HandleFunc("/api/get-person-by-cedula", getPersonByCedulaHandler)
	http.HandleFunc("/galeria", galleryHandler)
	http.HandleFunc("/upload", uploadHandler)
	http.HandleFunc("/delete/", deleteHandler)
	http.HandleFunc("/api/tree-data", getTreeData)
	http.HandleFunc("/api/get-people", getPeopleInHouse)
	http.HandleFunc("/api/pdf/export", exportToPDF)
	http.HandleFunc("/api/excel/export", exportToExcel)
	http.HandleFunc("/api/update-excel", updateExcelData)
	http.HandleFunc("/api/excel/columns", getColumns)
	http.HandleFunc("/api/excel", getData)
	http.HandleFunc("/editar-hogar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/editar_hogar.html")
	})
	http.HandleFunc("/api/get-household-details", getHouseholdDetails)
	http.HandleFunc("/api/add-household", addHouseholdData)
	http.HandleFunc("/agregar-hogar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/editar_hogar.html")
	})

	http.HandleFunc("/historia", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/historia.html")
	})

	http.HandleFunc("/listado_votantes", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/listado_votantes.html")
	})
	// -------------------------------

	fs := http.FileServer(http.Dir("assets"))
	http.Handle("/assets/", http.StripPrefix("/assets/", fs))

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/index.html")
	})
	http.HandleFunc("/base_de_datos", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/Base_de_Datos.html")
	})
	http.HandleFunc("/calendario", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/calendario.html")
	})
	http.HandleFunc("/comunidades", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/comunidades.html")
	})

	go func() {
		fmt.Println("Servidor corriendo en http://localhost:8080")
		if err := http.ListenAndServe(":8080", nil); err != nil {
			fmt.Println("Error:", err)
		}
	}()

	time.Sleep(500 * time.Millisecond)
	openBrowser("http://localhost:8080")
	select {}
}

//------------------- MANEJO DE LA GALERIA -------------------------

func galleryHandler(w http.ResponseWriter, r *http.Request) {
	files, err := ioutil.ReadDir(uploadDir)
	if err != nil {
		http.Error(w, "Error al leer imágenes", http.StatusInternalServerError)
		return
	}

	var images []string
	for _, file := range files {
		if !file.IsDir() {
			images = append(images, file.Name())
		}
	}

	tmpl := template.Must(template.ParseFiles("paginas/galeria.html"))
	data := GalleryData{
		Images:   images,
		Messages: []string{},
	}
	tmpl.Execute(w, data)
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Redirect(w, r, "/galeria", http.StatusSeeOther)
		return
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "Error al subir archivo", http.StatusBadRequest)
		return
	}
	defer file.Close()

	filename := filepath.Base(header.Filename)
	ext := strings.ToLower(filepath.Ext(filename))
	if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
		http.Error(w, "Formato no permitido", http.StatusBadRequest)
		return
	}

	dst, err := os.Create(filepath.Join(uploadDir, filename))
	if err != nil {
		http.Error(w, "Error al guardar imagen", http.StatusInternalServerError)
		return
	}
	defer dst.Close()

	io.Copy(dst, file)
	http.Redirect(w, r, "/galeria", http.StatusSeeOther)
}

func deleteHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Redirect(w, r, "/galeria", http.StatusSeeOther)
		return
	}

	filename := strings.TrimPrefix(r.URL.Path, "/delete/")
	filepath := filepath.Join(uploadDir, filename)

	if err := os.Remove(filepath); err != nil {
		http.Error(w, "No se pudo eliminar la imagen", http.StatusInternalServerError)
		return
	}

	http.Redirect(w, r, "/galeria", http.StatusSeeOther)
}

// getTreeData
func getTreeData(w http.ResponseWriter, r *http.Request) {
	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		fmt.Println("Error al abrir el archivo Excel:", err) // LOG
		http.Error(w, "no se pudo abrir el Excel: "+err.Error(), http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		fmt.Println("Error: La hoja está vacía o no se pudo leer.") // LOG
		http.Error(w, "sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comunidadIdx, torreIdx, casaIdx := -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		switch headerClean {
		case "comunidad":
			comunidadIdx = i
		case "torre":
			torreIdx = i
		case "casa o apto", "casa", "apto":
			casaIdx = i
		}
	}

	fmt.Println("--- Depurando la carga del árbol ---")
	fmt.Printf("Índice encontrado para 'Comunidad': %d\n", comunidadIdx)
	fmt.Printf("Índice encontrado para 'Torre': %d\n", torreIdx)
	fmt.Printf("Índice encontrado para 'Casa/Apto': %d\n", casaIdx)
	fmt.Println("------------------------------------")

	if comunidadIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		errorMsg := "No se encontraron todas las columnas requeridas. Revisa que tu Excel tenga cabeceras llamadas 'COMUNIDAD', 'TORRE' y 'CASA O APTO'."
		fmt.Println(errorMsg) // LOG
		http.Error(w, errorMsg, http.StatusInternalServerError)
		return
	}

	tree := make(map[string]map[string]map[string]struct{})
	for _, row := range rows[1:] {
		if len(row) <= comunidadIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}
		comunidad := strings.TrimSpace(row[comunidadIdx])
		torre := strings.TrimSpace(row[torreIdx])
		casa := strings.TrimSpace(row[casaIdx])
		if comunidad == "" || torre == "" || casa == "" {
			continue
		}
		if _, ok := tree[comunidad]; !ok {
			tree[comunidad] = make(map[string]map[string]struct{})
		}
		if _, ok := tree[comunidad][torre]; !ok {
			tree[comunidad][torre] = make(map[string]struct{})
		}
		tree[comunidad][torre][casa] = struct{}{}
	}

	var result []*TreeNode
	for comName, torres := range tree {
		comNode := &TreeNode{Text: comName, Type: "comunidad"}
		comNode.State.Opened = false

		for torreName, casas := range torres {
			torreNode := &TreeNode{Text: "Torre " + torreName, Type: "torre"}
			torreNode.State.Opened = false

			for casaName := range casas {
				casaNode := &TreeNode{Text: "Casa/Apto " + casaName, Type: "casa"}

				torreNode.Children = append(torreNode.Children, casaNode)
			}
			comNode.Children = append(comNode.Children, torreNode)
		}
		result = append(result, comNode)
	}

	fmt.Printf("Se encontraron %d comunidades para el árbol.\n", len(result)) // LOG

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(result)
}

// getPeopleInHouse
func getPeopleInHouse(w http.ResponseWriter, r *http.Request) {
	comunidad := r.URL.Query().Get("comunidad")
	torre := r.URL.Query().Get("torre")
	casa := r.URL.Query().Get("casa")

	if comunidad == "" || torre == "" || casa == "" {
		http.Error(w, "Faltan parámetros: comunidad, torre y casa son requeridos", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx, parentescoIdx, nombresIdx, docIdx := -1, -1, -1, -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		switch headerClean {
		case "comunidad":
			comIdx = i
		case "torre":
			torreIdx = i
		case "casa o apto", "casa", "apto":
			casaIdx = i
		case "parentesco":
			parentescoIdx = i
		case "nombre completo":
			nombresIdx = i
		case "cedula de identidad":
			docIdx = i
		}
	}

	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 || parentescoIdx == -1 || nombresIdx == -1 || docIdx == -1 {
		http.Error(w, "No se encontraron todas las columnas requeridas en el Excel", http.StatusInternalServerError)
		return
	}

	var people []Person
	for _, row := range rows[1:] {
		if len(row) > comIdx && len(row) > torreIdx && len(row) > casaIdx {
			if row[comIdx] == comunidad && row[torreIdx] == torre && row[casaIdx] == casa {
				person := Person{}
				if parentescoIdx < len(row) {
					person.Parentesco = row[parentescoIdx]
				}
				if nombresIdx < len(row) {
					person.Nombres = row[nombresIdx]
				}
				if docIdx < len(row) {
					person.Documento = row[docIdx]
				}
				people = append(people, person)
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(people)
}

// getHouseholdDetails
func getHouseholdDetails(w http.ResponseWriter, r *http.Request) {
	comunidad := r.URL.Query().Get("comunidad")
	torre := r.URL.Query().Get("torre")
	casa := r.URL.Query().Get("casa")

	if comunidad == "" || torre == "" || casa == "" {
		http.Error(w, "Faltan parámetros", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "Sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx := -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		if headerClean == "comunidad" {
			comIdx = i
		}
		if headerClean == "torre" {
			torreIdx = i
		}
		if headerClean == "casa o apto" || headerClean == "casa" {
			casaIdx = i
		}
	}

	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas clave no encontradas", http.StatusInternalServerError)
		return
	}

	var householdData []map[string]string
	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) > comIdx && len(row) > torreIdx && len(row) > casaIdx {
			if row[comIdx] == comunidad && row[torreIdx] == torre && row[casaIdx] == casa {
				personData := make(map[string]string)
				personData["__row"] = strconv.Itoa(i + 1)
				for j, header := range headers {
					cleanHeader := strings.TrimSpace(header)
					if j < len(row) {
						personData[cleanHeader] = row[j]
					} else {
						personData[cleanHeader] = ""
					}
				}
				householdData = append(householdData, personData)
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(householdData)
}

// addHouseholdData
func addHouseholdData(w http.ResponseWriter, r *http.Request) {
	var req struct {
		Datos []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]
	nextRow := len(rows) + 1

	headerMap := make(map[string]int)
	for i, h := range headers {
		headerMap[strings.TrimSpace(h)] = i // Usar clave limpia
	}

	for _, persona := range req.Datos {
		for key, val := range persona {
			if colIndex, ok := headerMap[key]; ok { // 'key' ya viene limpia del frontend
				cell, _ := excelize.CoordinatesToCellName(colIndex+1, nextRow)
				f.SetCellValue(PRIMERA_HOJA, cell, val)
			}
		}
		nextRow++
	}

	if err := f.Save(); err != nil {
		http.Error(w, "No se guardó el Excel", http.StatusInternalServerError)
		return
	}
	w.WriteHeader(http.StatusOK)
}

// Google
func handleGoogleLogin(w http.ResponseWriter, r *http.Request) {
	url := googleOauthConfig.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	http.Redirect(w, r, url, http.StatusTemporaryRedirect)
}
func handleGoogleCallback(w http.ResponseWriter, r *http.Request) {
	ctx := context.Background()
	code := r.URL.Query().Get("code")

	// 1. Intercambiar código por Token
	token, err := googleOauthConfig.Exchange(ctx, code)
	if err != nil {
		http.Error(w, "Fallo al intercambiar token", http.StatusInternalServerError)
		return
	}

	// 2. Obtener información del usuario desde Google
	client := googleOauthConfig.Client(ctx, token)
	resp, err := client.Get("https://www.googleapis.com/oauth2/v2/userinfo")
	if err != nil {
		http.Error(w, "Fallo al obtener datos de usuario", http.StatusInternalServerError)
		return
	}
	defer resp.Body.Close()

	var userInfo struct {
		Email   string `json:"email"`
		Name    string `json:"name"`
		Picture string `json:"picture"`
	}
	json.NewDecoder(resp.Body).Decode(&userInfo)

	// 3. Verificar si el usuario está autorizado
	role, authorized := authorizedUsers[userInfo.Email]
	if !authorized {
		// Usuario no está en la lista de los 3 permitidos
		http.Redirect(w, r, "/?error=unauthorized", http.StatusTemporaryRedirect)
		return
	}

	// 4. Crear la sesión activa
	activeSession = &UserSession{
		Email:   userInfo.Email,
		Name:    userInfo.Name,
		Picture: userInfo.Picture,
		Role:    role,
		Token:   token, // <-- Agregamos el token aquí para poder usar Drive luego
	}

	// 5. Inicializar el servicio de Drive para este usuario
	// (Opcional: aquí podrías buscar el archivo en Drive para tener el remoteFileID listo)
	fmt.Printf("Usuario logueado: %s con rol %s\n", activeSession.Email, activeSession.Role)

	http.Redirect(w, r, "/", http.StatusSeeOther)
}

// Estructura para informar al frontend del estado de sincronización
type SyncStatus struct {
	NeedsPull bool   `json:"needsPull"` // Hay algo nuevo en Drive
	NeedsPush bool   `json:"needsPush"` // Hay algo nuevo local
	LastSync  string `json:"lastSync"`
}

func getSyncStatus(w http.ResponseWriter, r *http.Request) {
	if activeSession == nil {
		http.Error(w, "No logueado", http.StatusUnauthorized)
		return
	}

	// 1. Obtener info del archivo local
	localInfo, err := os.Stat(EXCEL_FILE)
	if err != nil {
		http.Error(w, "Archivo local no encontrado", 500)
		return
	}
	localTime := localInfo.ModTime()

	// 2. Lógica de comparación básica (Por ahora simulamos que no necesita Pull)
	// Para evitar el error de "localTime declared and not used", imprimimos o comparamos:
	fmt.Printf("Verificando sincronización. Fecha local: %v\n", localTime)

	status := SyncStatus{
		NeedsPull: false,
		NeedsPush: true, // Asumimos push por defecto para probar
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(status)
}

// Obtiene el cliente de Drive usando el token de la sesión activa
func getDriveService() (*drive.Service, error) {
	if activeSession == nil || activeSession.Token == nil {
		return nil, fmt.Errorf("no hay sesión activa")
	}
	ctx := context.Background()
	return drive.NewService(ctx, option.WithTokenSource(googleOauthConfig.TokenSource(ctx, activeSession.Token)))
}

// Busca el archivo Excel en Drive o lo crea si no existe
func getOrCreateDriveFile(srv *drive.Service) (string, error) {
	query := fmt.Sprintf("name = '%s' and trashed = false", EXCEL_FILE)
	list, err := srv.Files.List().Q(query).Do()
	if err != nil {
		return "", err
	}

	if len(list.Files) > 0 {
		return list.Files[0].Id, nil
	}

	// Si no existe, lo creamos subiendo la copia local actual
	f, _ := os.Open(EXCEL_FILE)
	defer f.Close()
	driveFile, err := srv.Files.Create(&drive.File{Name: EXCEL_FILE}).Media(f).Do()
	if err != nil {
		return "", err
	}
	return driveFile.Id, nil
}

// Handler para el PULL (Bajar de Drive a Local)
func handlePull(w http.ResponseWriter, r *http.Request) {
	srv, err := getDriveService()
	if err != nil {
		http.Error(w, err.Error(), 401)
		return
	}

	fileID, err := getOrCreateDriveFile(srv)
	if err != nil {
		http.Error(w, err.Error(), 500)
		return
	}

	resp, err := srv.Files.Get(fileID).Download()
	if err != nil {
		http.Error(w, err.Error(), 500)
		return
	}
	defer resp.Body.Close()

	out, _ := os.Create(EXCEL_FILE)
	defer out.Close()
	io.Copy(out, resp.Body)

	fmt.Println("Sincronización PULL completada.")
	w.WriteHeader(http.StatusOK)
}

// Handler para el PUSH (Subir de Local a Drive)
func handlePush(w http.ResponseWriter, r *http.Request) {
	srv, err := getDriveService()
	if err != nil {
		http.Error(w, err.Error(), 401)
		return
	}

	fileID, err := getOrCreateDriveFile(srv)
	if err != nil {
		http.Error(w, err.Error(), 500)
		return
	}

	f, _ := os.Open(EXCEL_FILE)
	defer f.Close()

	_, err = srv.Files.Update(fileID, &drive.File{}).Media(f).Do()
	if err != nil {
		http.Error(w, err.Error(), 500)
		return
	}

	fmt.Println("Sincronización PUSH completada.")
	w.WriteHeader(http.StatusOK)
}

// Handler para que el frontend sepa quién está logueado
func getUserStatus(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	if activeSession == nil {
		json.NewEncoder(w).Encode(map[string]interface{}{"logged": false})
		return
	}

	// Enviamos los datos al frontend
	json.NewEncoder(w).Encode(map[string]interface{}{
		"logged":  true,
		"email":   activeSession.Email,
		"name":    activeSession.Name,
		"picture": activeSession.Picture,
		"role":    activeSession.Role,
	})
}
