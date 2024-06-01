package com.example.bilancio

import android.content.Context
import android.content.Intent
import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.PopupMenu
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class indici : AppCompatActivity() {
    var context: Context = this

    // Dichiarazione degli EditText come membri della classe
    var capitale_di_terzi_ind_globale: EditText? = null
    var patrimonio_netto_ind_globale: EditText? = null
    var banca_ind_strutturale: EditText? = null
    var quotamutuipassivi_ind_strutturale: EditText? = null
    var mutuipassivi_ind_strutturale: EditText? = null
    var patrimonio_netto_ind_strutturale: EditText? = null
    var pas_brevetermine_ind_esigibilità: EditText? = null
    var patrimonio_netto_ind_esigibilità: EditText? = null
    var attivo_fisso_netto_ind_rigidità_attivo: EditText? = null
    var attivo_totale_netto_ind_rigidità_attivo: EditText? = null
    var attivo_a_breve_ind_elasticità_attivo: EditText? = null
    var attivo_totale_netto_ind_elasticità_attivo: EditText? = null
    var patrimonio_netto_ind_grado_di_copertura_1: EditText? = null
    var attivo_fisso_netto_ind_grado_di_copertura_1: EditText? = null
    var patrimonio_netto_ind_grado_di_copertura_2: EditText? = null
    var passivo_medio_lungo_termine_ind_grado_di_copertura_2: EditText? = null
    var attivo_fisso_netto_ind_grado_di_copertura_2: EditText? = null
    var fondo_ammortamento_ind_grado_di_ammortamento: EditText? = null
    var immobilizzazioni_materiali_ind_grado_di_ammortamento: EditText? = null
    var roa_ind_roi: EditText? = null
    var attivo_totale_netto_ind_roi: EditText? = null
    var roa_ind_ros: EditText? = null
    var ricavi_netti_vendita_ind_ros: EditText? = null
    var reddito_netto_utile_vendita_ind_roe: EditText? = null
    var patrimonio_netto_ind_roe: EditText? = null
    var ricavi_netti_vendita_ind_turnover: EditText? = null
    var attivo_totale_netto_ind_turnover: EditText? = null
    var cassa_ind_liquidità_1: EditText? = null
    var banca_ind_liquidità_1: EditText? = null
    var crediti_ind_liquidità_1: EditText? = null
    var passività_ind_liquidità_1: EditText? = null
    var attività_ind_liquidità_2: EditText? = null
    var passività_ind_liquidità_2: EditText? = null
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

       
        setContentView(R.layout.activity_indici)

        // Inizializzazione degli EditText
        capitale_di_terzi_ind_globale = findViewById(R.id.capitale_di_terzi_ind_globale)
        patrimonio_netto_ind_globale = findViewById(R.id.patrimonio_netto_ind_globale)
        banca_ind_strutturale = findViewById(R.id.banca_ind_strutturale)
        quotamutuipassivi_ind_strutturale = findViewById(R.id.quotamutuipassivi_ind_strutturale)
        mutuipassivi_ind_strutturale = findViewById(R.id.mutuipassivi_ind_strutturale)
        patrimonio_netto_ind_strutturale = findViewById(R.id.patrimonio_netto_ind_strutturale)
        pas_brevetermine_ind_esigibilità = findViewById(R.id.pas_brevetermine_ind_esigibilità)
        patrimonio_netto_ind_esigibilità = findViewById(R.id.patrimonio_netto_ind_esigibilità)
        attivo_fisso_netto_ind_rigidità_attivo =
            findViewById(R.id.attivo_fisso_netto_ind_rigidità_attivo)
        attivo_totale_netto_ind_rigidità_attivo =
            findViewById(R.id.attivo_totale_netto_ind_rigidità_attivo)
        attivo_a_breve_ind_elasticità_attivo =
            findViewById(R.id.attivo_a_breve_ind_elasticità_attivo)
        attivo_totale_netto_ind_elasticità_attivo =
            findViewById(R.id.attivo_totale_netto_ind_elasticità_attivo)
        patrimonio_netto_ind_grado_di_copertura_1 =
            findViewById(R.id.patrimonio_netto_ind_grado_di_copertura_1)
        attivo_fisso_netto_ind_grado_di_copertura_1 =
            findViewById(R.id.attivo_fisso_netto_ind_grado_di_copertura_1)
        patrimonio_netto_ind_grado_di_copertura_2 =
            findViewById(R.id.patrimonio_netto_ind_grado_di_copertura_2)
        passivo_medio_lungo_termine_ind_grado_di_copertura_2 =
            findViewById(R.id.passivo_medio_lungo_termine_ind_grado_di_copertura_2)
        attivo_fisso_netto_ind_grado_di_copertura_2 =
            findViewById(R.id.attivo_fisso_netto_ind_grado_di_copertura_2)
        fondo_ammortamento_ind_grado_di_ammortamento =
            findViewById(R.id.fondo_ammortamento_ind_grado_di_ammortamento)
        immobilizzazioni_materiali_ind_grado_di_ammortamento =
            findViewById(R.id.immobilizzazioni_materiali_ind_grado_di_ammortamento)
        roa_ind_roi = findViewById(R.id.roa_ind_roi)
        attivo_totale_netto_ind_roi = findViewById(R.id.attivo_totale_netto_ind_roi)
        roa_ind_ros = findViewById(R.id.roa_ind_ros)
        ricavi_netti_vendita_ind_ros = findViewById(R.id.ricavi_netti_vendita_ind_ros)
        reddito_netto_utile_vendita_ind_roe = findViewById(R.id.reddito_netto_utile_vendita_ind_roe)
        patrimonio_netto_ind_roe = findViewById(R.id.patrimonio_netto_ind_roe)
        ricavi_netti_vendita_ind_turnover = findViewById(R.id.ricavi_netti_vendita_ind_turnover)
        attivo_totale_netto_ind_turnover = findViewById(R.id.attivo_totale_netto_ind_turnover)
        cassa_ind_liquidità_1 = findViewById(R.id.cassa_ind_liquidità_1)
        banca_ind_liquidità_1 = findViewById(R.id.banca_ind_liquidità_1)
        crediti_ind_liquidità_1 = findViewById(R.id.crediti_ind_liquidità_1)
        passività_ind_liquidità_1 = findViewById(R.id.passività_ind_liquidità_1)
        attività_ind_liquidità_2 = findViewById(R.id.attività_ind_liquidità_2)
        passività_ind_liquidità_2 = findViewById(R.id.passività_ind_liquidità_2)
        val salvaa = findViewById<Button>(R.id.salvaindici)
        salvaa.setOnClickListener { // Crea una nuova istanza di IndiciDiBilancioSaver
            val fileSaver = IndiciDiBilancioSaver(context)

            // Crea un nuovo workbook e imposta i valori
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("INDICI DI SOLIDITÀ")

            // Ciclo per creare le righe con i valori
            for (i in 0..34) {
                val row = sheet.createRow(i)
                // Creazione delle 5 colonne per ogni riga
                for (j in 0..3) {
                    val cell = row.createCell(j)
                    // Imposta i valori in base alle condizioni richieste
                    if (j == 0) {
                        cell.setCellValue("") // Colonna A
                    } else if (j == 1) {
                        cell.setCellValue("") // Colonna B
                    } else if (j == 2) {
                        cell.setCellValue("") // Colonna C
                    } else if (j == 3) {
                        cell.setCellValue("") // Colonna D
                    }
                }
            }

            // Ottieni l'oggetto CellStyle per le celle unite
            val mergedCellStyle = workbook.createCellStyle()
            mergedCellStyle.borderBottom = BorderStyle.THIN // Bordo inferiore
            mergedCellStyle.borderTop = BorderStyle.THIN // Bordo superiore
            mergedCellStyle.borderLeft = BorderStyle.THIN // Bordo sinistro
            mergedCellStyle.borderRight = BorderStyle.THIN // Bordo destro

            // Stile orizzontale
            val horizontalstyle = workbook.createCellStyle()
            // Imposta l'allineamento del testo al centro
            horizontalstyle.alignment = HorizontalAlignment.CENTER

            // Bordo piccolo
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet)
            RegionUtil.setBorderLeft(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 1), sheet)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(2, 2, 3, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 0), sheet)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(3, 3, 0, 0), sheet)


            //Bordo medio alto
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(11, 11, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(15, 15, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(19, 19, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(23, 23, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(27, 27, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(31, 31, 0, 3), sheet)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(35, 35, 0, 3), sheet)

            //Bordo medio basso
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(3, 3, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(11, 11, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(15, 15, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(19, 19, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(23, 23, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(27, 27, 0, 3), sheet)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(31, 31, 0, 3), sheet)

            // Imposta la larghezza delle colonne
            sheet.setColumnWidth(0, 5000)
            sheet.setColumnWidth(1, 5000)
            sheet.setColumnWidth(2, 5000)
            sheet.setColumnWidth(3, 5000)


            // Unisci celle
            sheet.addMergedRegion(CellRangeAddress(0, 0, 0, 3))
            sheet.addMergedRegion(CellRangeAddress(2, 2, 0, 2))
            sheet.addMergedRegion(CellRangeAddress(4, 4, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(8, 8, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(12, 12, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(16, 16, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(20, 20, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(24, 24, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(28, 28, 0, 1))
            sheet.addMergedRegion(CellRangeAddress(32, 32, 0, 1))

            // Stile per il testo in grassetto e colore rosso
            val style = workbook.createCellStyle()
            val font: Font = workbook.createFont()
            font.bold = true
            font.color = IndexedColors.RED.getIndex()
            style.setFont(font)

            // Popolamento celle
            val rowA1 = sheet.getRow(0)
            val cellA1 = rowA1.createCell(0)
            cellA1.setCellValue("INDICI DI BILANCIO - SOLIDITÀ")
            cellA1.setCellStyle(horizontalstyle)
            val rowA3 = sheet.getRow(2)
            val cellA3 = rowA3.createCell(0)
            cellA3.setCellValue("INDICI DI SOLIDITÀ")
            cellA3.setCellStyle(horizontalstyle)
            val rowA5 = sheet.getRow(4)
            val cellA5 = rowA5.createCell(0)
            cellA5.setCellStyle(style)
            cellA5.setCellValue("INDEBITAMENTO GLOBALE")
            val rowD5 = sheet.getRow(4)
            val cellD5 = rowD5.createCell(3)
            val rowC5 = sheet.getRow(4)
            val cellC5 = rowC5.createCell(2)
            cellC5.setCellValue("CAPITALE DI TERZI")
            val rowC6 = sheet.getRow(5)
            val cellC6 = rowC6.createCell(2)
            cellC6.setCellValue("PATRIMONIO NETTO")
            val rowD6 = sheet.getRow(5)
            val cellD6 = rowD6.createCell(3)
            val rowA7 = sheet.getRow(6)
            val cellA7 = rowA7.createCell(0)
            cellA7.setCellValue("TOTALE")
            val rowD7 = sheet.getRow(6)
            val cellD7 = rowD7.createCell(3)
            if (capitale_di_terzi_ind_globale!!.getText()
                    .toString().length != 0 && patrimonio_netto_ind_globale!!.getText()
                    .toString().length != 0
            ) {
                val capitalediterziindglobale =
                    capitale_di_terzi_ind_globale!!.getText().toString().toDouble()
                val patrimonionettoindglobale =
                    patrimonio_netto_ind_globale!!.getText().toString().toDouble()
                cellD5.setCellValue(capitalediterziindglobale)
                cellD6.setCellValue(patrimonionettoindglobale)
                cellD7.cellFormula = "D5/D6"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet)
            }
            val rowA9 = sheet.getRow(8)
            val cellA9 = rowA9.createCell(0)
            cellA9.setCellValue("INDEBITAMENTO STRUTTURALE")
            val rowD9 = sheet.getRow(8)
            val cellD9 = rowD9.createCell(3)
            val rowC9 = sheet.getRow(8)
            val cellC9 = rowC9.createCell(2)
            cellA9.setCellStyle(style)
            cellC9.setCellValue("PASSIVITÀ FINANZIARIE")
            val rowC10 = sheet.getRow(9)
            val cellC10 = rowC10.createCell(2)
            cellC10.setCellValue("PATRIMONIO NETTO")
            val rowD10 = sheet.getRow(9)
            val cellD10 = rowD10.createCell(3)
            val rowA11 = sheet.getRow(10)
            val cellA11 = rowA11.createCell(0)
            cellA11.setCellValue("TOTALE")
            val rowD11 = sheet.getRow(10)
            val cellD11 = rowD11.createCell(3)
            if (banca_ind_strutturale!!.getText()
                    .toString().length != 0 || quotamutuipassivi_ind_strutturale!!.getText()
                    .toString().length != 0 || mutuipassivi_ind_strutturale!!.getText()
                    .toString().length != 0 && patrimonio_netto_ind_strutturale!!.getText()
                    .toString().length != 0
            ) {
                val bancaccpassivostrutturale =
                    banca_ind_strutturale!!.getText().toString().toDouble()
                val quotamutuipassivistrutturale =
                    quotamutuipassivi_ind_strutturale!!.getText().toString().toDouble()
                val mutuipassiviindstrutturale =
                    mutuipassivi_ind_strutturale!!.getText().toString().toDouble()
                val patrimonionettoindstrutturale =
                    patrimonio_netto_ind_strutturale!!.getText().toString().toDouble()
                cellD9.setCellValue(bancaccpassivostrutturale + quotamutuipassivistrutturale + mutuipassiviindstrutturale)
                cellD10.setCellValue(patrimonionettoindstrutturale)
                cellD11.cellFormula = "D9/D10"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(10, 10, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(10, 10, 3, 3), sheet)
            }
            val rowA13 = sheet.getRow(12)
            val cellA13 = rowA13.createCell(0)
            cellA13.setCellStyle(style)
            cellA13.setCellValue("ESIGIBILITÀ DEL PASSIVO")
            val rowD13 = sheet.getRow(12)
            val cellD13 = rowD13.createCell(3)
            val rowC13 = sheet.getRow(12)
            val cellC13 = rowC13.createCell(2)
            cellC13.setCellValue("PASSIVITÀ A BREVE TERMINE")
            val rowC14 = sheet.getRow(13)
            val cellC14 = rowC14.createCell(2)
            cellC14.setCellValue("CAPITALE DI TERZI")
            val rowD14 = sheet.getRow(13)
            val cellD14 = rowD14.createCell(3)
            val rowA15 = sheet.getRow(14)
            val cellA15 = rowA15.createCell(0)
            cellA15.setCellValue("TOTALE")
            val rowD15 = sheet.getRow(14)
            val cellD15 = rowD15.createCell(3)
            if (pas_brevetermine_ind_esigibilità!!.getText()
                    .toString().length != 0 && patrimonio_netto_ind_esigibilità!!.getText()
                    .toString().length != 0
            ) {
                val passivitaabreveindpassivo =
                    pas_brevetermine_ind_esigibilità!!.getText().toString().toDouble()
                val patrimonionettoindpassivo =
                    patrimonio_netto_ind_esigibilità!!.getText().toString().toDouble()
                cellD13.setCellValue(passivitaabreveindpassivo)
                cellD14.setCellValue(patrimonionettoindpassivo)
                cellD15.cellFormula = "D13/D14"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(14, 14, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(14, 14, 3, 3), sheet)
            }
            val rowA17 = sheet.getRow(16)
            val cellA17 = rowA17.createCell(0)
            cellA17.setCellStyle(style)
            cellA17.setCellValue("GRADO DI RIGIDITÀ DELL'ATTIVO")
            val rowD17 = sheet.getRow(16)
            val cellD17 = rowD17.createCell(3)
            val rowC17 = sheet.getRow(16)
            val cellC17 = rowC17.createCell(2)
            cellC17.setCellValue("ATTIVO FISSO NETTO")
            val rowC18 = sheet.getRow(17)
            val cellC18 = rowC18.createCell(2)
            cellC18.setCellValue("ATTIVO TOTALE NETTO")
            val rowD18 = sheet.getRow(17)
            val cellD18 = rowD18.createCell(3)
            val rowA19 = sheet.getRow(18)
            val cellA19 = rowA19.createCell(0)
            cellA19.setCellValue("TOTALE")
            val rowD19 = sheet.getRow(18)
            val cellD19 = rowD19.createCell(3)
            if (attivo_fisso_netto_ind_rigidità_attivo!!.getText()
                    .toString().length != 0 && attivo_totale_netto_ind_rigidità_attivo!!.getText()
                    .toString().length != 0
            ) {
                val attivofissonettorigidità =
                    attivo_fisso_netto_ind_rigidità_attivo!!.getText().toString().toDouble()
                val attivototalenettorigidità =
                    attivo_totale_netto_ind_rigidità_attivo!!.getText().toString().toDouble()
                cellD17.setCellValue(attivofissonettorigidità)
                cellD18.setCellValue(attivototalenettorigidità)
                cellD19.cellFormula = "D17/D18"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(18, 18, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(18, 18, 3, 3), sheet)
            }
            val rowA21 = sheet.getRow(20)
            val cellA21 = rowA21.createCell(0)
            cellA21.setCellStyle(style)
            cellA21.setCellValue("GRADO DI ELASTICITÀ DELL'ATTIVO")
            val rowD21 = sheet.getRow(20)
            val cellD21 = rowD21.createCell(3)
            val rowC21 = sheet.getRow(20)
            val cellC21 = rowC21.createCell(2)
            cellC21.setCellValue("ATTIVO A BREVE")
            val rowC22 = sheet.getRow(21)
            val cellC22 = rowC22.createCell(2)
            cellC22.setCellValue("ATTIVO TOTALE NETTO")
            val rowD22 = sheet.getRow(21)
            val cellD22 = rowD22.createCell(3)
            val rowA23 = sheet.getRow(22)
            val cellA23 = rowA23.createCell(0)
            cellA23.setCellValue("TOTALE")
            val rowD23 = sheet.getRow(22)
            val cellD23 = rowD23.createCell(3)
            if (attivo_a_breve_ind_elasticità_attivo!!.getText()
                    .toString().length != 0 && attivo_totale_netto_ind_elasticità_attivo!!.getText()
                    .toString().length != 0
            ) {
                val attivoabreveelasticità =
                    attivo_a_breve_ind_elasticità_attivo!!.getText().toString().toDouble()
                val attivototalenettoelasticità =
                    attivo_totale_netto_ind_elasticità_attivo!!.getText().toString().toDouble()
                cellD21.setCellValue(attivoabreveelasticità)
                cellD22.setCellValue(attivototalenettoelasticità)
                cellD23.cellFormula = "D21/D22"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(22, 22, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(22, 22, 3, 3), sheet)
            }
            val rowA25 = sheet.getRow(24)
            val cellA25 = rowA25.createCell(0)
            cellA25.setCellStyle(style)
            cellA25.setCellValue("GRADO DI COPERTURA DELL'ATTIVO (I) LIVELLO")
            val rowD25 = sheet.getRow(24)
            val cellD25 = rowD25.createCell(3)
            val rowC25 = sheet.getRow(24)
            val cellC25 = rowC25.createCell(2)
            cellC25.setCellValue("PATRIMONIO NETTO")
            val rowC26 = sheet.getRow(25)
            val cellC26 = rowC26.createCell(2)
            cellC26.setCellValue("ATTIVO FISSO NETTO")
            val rowD26 = sheet.getRow(25)
            val cellD26 = rowD26.createCell(3)
            val rowA27 = sheet.getRow(26)
            val cellA27 = rowA27.createCell(0)
            cellA27.setCellValue("TOTALE")
            val rowD27 = sheet.getRow(26)
            val cellD27 = rowD27.createCell(3)
            if (patrimonio_netto_ind_grado_di_copertura_1!!.getText()
                    .toString().length != 0 && attivo_fisso_netto_ind_grado_di_copertura_1!!.getText()
                    .toString().length != 0
            ) {
                val patrimonionettocopertura =
                    patrimonio_netto_ind_grado_di_copertura_1!!.getText().toString().toDouble()
                val attivofissonettocopertura =
                    attivo_fisso_netto_ind_grado_di_copertura_1!!.getText().toString().toDouble()
                cellD25.setCellValue(patrimonionettocopertura)
                cellD26.setCellValue(attivofissonettocopertura)
                cellD27.cellFormula = "D25/D26"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(26, 26, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(26, 26, 3, 3), sheet)
            }
            val rowA29 = sheet.getRow(28)
            val cellA29 = rowA29.createCell(0)
            cellA29.setCellStyle(style)
            cellA29.setCellValue("GRADO DI COPERTURA DELL'ATTIVO (II) LIVELLO")
            val rowD29 = sheet.getRow(28)
            val cellD29 = rowD29.createCell(3)
            val rowC29 = sheet.getRow(28)
            val cellC29 = rowC29.createCell(2)
            cellC29.setCellValue(
                """
    PATRIMONIO NETTO + 
    PASSIVO A MEDIO/LUNGO
    TERMINE
    """.trimIndent()
            )
            val rowC30 = sheet.getRow(29)
            val cellC30 = rowC30.createCell(2)
            cellC30.setCellValue("ATTIVO FISSO NETTO")
            val rowD30 = sheet.getRow(29)
            val cellD30 = rowD30.createCell(3)
            val rowA31 = sheet.getRow(30)
            val cellA31 = rowA31.createCell(0)
            cellA31.setCellValue("TOTALE")
            val rowD31 = sheet.getRow(30)
            val cellD31 = rowD31.createCell(3)
            if (patrimonio_netto_ind_grado_di_copertura_2!!.getText()
                    .toString().length != 0 || passivo_medio_lungo_termine_ind_grado_di_copertura_2!!.getText()
                    .toString().length != 0 && attivo_fisso_netto_ind_grado_di_copertura_2!!.getText()
                    .toString().length != 0
            ) {
                val patrimonionettocopertura2 =
                    patrimonio_netto_ind_grado_di_copertura_2!!.getText().toString().toDouble()
                val passivomediolungotermine =
                    passivo_medio_lungo_termine_ind_grado_di_copertura_2!!.getText().toString()
                        .toDouble()
                val attivofissonettocopertura2 =
                    attivo_fisso_netto_ind_grado_di_copertura_2!!.getText().toString().toDouble()
                cellD29.setCellValue(patrimonionettocopertura2 + passivomediolungotermine)
                cellD30.setCellValue(attivofissonettocopertura2)
                cellD31.cellFormula = "D29/D30"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(30, 30, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(30, 30, 3, 3), sheet)
            }
            val rowA33 = sheet.getRow(32)
            val cellA33 = rowA33.createCell(0)
            cellA33.setCellStyle(style)
            cellA33.setCellValue("GRADO DI AMMORTAMENTO")
            val rowD33 = sheet.getRow(32)
            val cellD33 = rowD33.createCell(3)
            val rowC33 = sheet.getRow(32)
            val cellC33 = rowC33.createCell(2)
            cellC33.setCellValue("FONDO AMMORTAMENTO")
            val rowC34 = sheet.getRow(33)
            val cellC34 = rowC34.createCell(2)
            cellC34.setCellValue("IMMOBILIZZAZIONI MATERIALI")
            val rowD34 = sheet.getRow(33)
            val cellD34 = rowD34.createCell(3)
            val rowA35 = sheet.getRow(34)
            val cellA35 = rowA35.createCell(0)
            cellA35.setCellValue("TOTALE")
            val rowD35 = sheet.getRow(34)
            val cellD35 = rowD35.createCell(3)
            if (fondo_ammortamento_ind_grado_di_ammortamento!!.getText()
                    .toString().length != 0 && immobilizzazioni_materiali_ind_grado_di_ammortamento!!.getText()
                    .toString().length != 0
            ) {
                val fondoammortamentogrado =
                    fondo_ammortamento_ind_grado_di_ammortamento!!.getText().toString().toDouble()
                val immobilizzazionimateriali =
                    immobilizzazioni_materiali_ind_grado_di_ammortamento!!.getText().toString()
                        .toDouble()
                cellD33.setCellValue(fondoammortamentogrado)
                cellD34.setCellValue(immobilizzazionimateriali)
                cellD35.cellFormula = "D33/D34"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(34, 34, 3, 3), sheet)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(34, 34, 3, 3), sheet)
            }

            // Bordo medio destro dopo aver aggiunto i valori
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(4, 5, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(8, 9, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(12, 13, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(16, 17, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(20, 21, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(24, 25, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(28, 29, 3, 3), sheet)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(32, 33, 3, 3), sheet)

            /*
    
    
                    FOGLIO 2
    
                     */
            val sheet2 = workbook.createSheet("INDICI DI REDDITIVITÀ")

            // Ciclo per creare le righe con i valori
            for (i in 0..18) {
                val row = sheet2.createRow(i)
                for (j in 0..3) {
                    val cell = row.createCell(j)
                }
            }
            // Imposta la larghezza delle colonne
            sheet2.setColumnWidth(0, 5000)
            sheet2.setColumnWidth(1, 5000)
            sheet2.setColumnWidth(2, 5000)
            sheet2.setColumnWidth(3, 5000)

            // Unisci celle
            sheet2.addMergedRegion(CellRangeAddress(0, 0, 0, 3))
            sheet2.addMergedRegion(CellRangeAddress(2, 2, 0, 2))
            sheet2.addMergedRegion(CellRangeAddress(4, 4, 0, 1))
            sheet2.addMergedRegion(CellRangeAddress(8, 8, 0, 1))
            sheet2.addMergedRegion(CellRangeAddress(12, 12, 0, 1))
            sheet2.addMergedRegion(CellRangeAddress(16, 16, 0, 1))

            // Bordo piccolo
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet2)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet2)
            RegionUtil.setBorderLeft(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet2)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet2)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 1), sheet2)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet2)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(2, 2, 3, 3), sheet2)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet2)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 0), sheet2)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(3, 3, 0, 0), sheet2)


            //Bordo medio alto
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet2)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(11, 11, 0, 3), sheet2)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(15, 15, 0, 3), sheet2)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(19, 19, 0, 3), sheet2)

            //Bordo medio basso
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(3, 3, 0, 3), sheet2)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet2)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(11, 11, 0, 3), sheet2)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(15, 15, 0, 3), sheet2)


            // Popolamento celle
            val rowAA1 = sheet2.getRow(0)
            val cellAA1 = rowAA1.createCell(0)
            cellAA1.setCellValue("INDICI DI BILANCIO - REDDITIVITÀ")
            cellAA1.setCellStyle(horizontalstyle)
            val rowAA3 = sheet2.getRow(2)
            val cellAA3 = rowAA3.createCell(0)
            cellAA3.setCellValue("INDICI DI REDDITIVITÀ")
            cellAA3.setCellStyle(horizontalstyle)
            val rowAA5 = sheet2.getRow(4)
            val cellAA5 = rowAA5.createCell(0)
            cellAA5.setCellStyle(style)
            cellAA5.setCellValue("ROI")
            val rowDD5 = sheet2.getRow(4)
            val cellDD5 = rowDD5.createCell(3)
            val rowCC5 = sheet2.getRow(4)
            val cellCC5 = rowCC5.createCell(2)
            cellCC5.setCellValue("REDDITO GESTIONE OPERATIVO")
            val rowCC6 = sheet2.getRow(5)
            val cellCC6 = rowCC6.createCell(2)
            cellCC6.setCellValue("ATTIVO TOTALE NETTO")
            val rowDD6 = sheet2.getRow(5)
            val cellDD6 = rowDD6.createCell(3)
            val rowAA7 = sheet2.getRow(6)
            val cellAA7 = rowAA7.createCell(0)
            cellAA7.setCellValue("TOTALE")
            val rowDD7 = sheet2.getRow(6)
            val cellDD7 = rowDD7.createCell(3)
            if (roa_ind_roi!!.getText()
                    .toString().length != 0 && attivo_totale_netto_ind_roi!!.getText()
                    .toString().length != 0
            ) {
                val redditogestioneoperativoroi = roa_ind_roi!!.getText().toString().toDouble()
                val attivototalenettoroi =
                    attivo_totale_netto_ind_roi!!.getText().toString().toDouble()
                cellDD5.setCellValue(redditogestioneoperativoroi)
                cellDD6.setCellValue(attivototalenettoroi)
                cellDD7.cellFormula = "D5/D6"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet2)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet2)
            }
            val rowAA9 = sheet2.getRow(8)
            val cellAA9 = rowAA9.createCell(0)
            cellAA9.setCellStyle(style)
            cellAA9.setCellValue("ROS")
            val rowDD9 = sheet2.getRow(8)
            val cellDD9 = rowDD9.createCell(3)
            val rowCC9 = sheet2.getRow(8)
            val cellCC9 = rowCC9.createCell(2)
            cellCC9.setCellValue("REDDITO GESTIONE OPERATIVO")
            val rowCC10 = sheet2.getRow(9)
            val cellCC10 = rowCC10.createCell(2)
            cellCC10.setCellValue("RICAVI NETTI DI VENDITA")
            val rowDD10 = sheet2.getRow(9)
            val cellDD10 = rowDD10.createCell(3)
            val rowAA11 = sheet2.getRow(10)
            val cellAA11 = rowAA11.createCell(0)
            cellAA11.setCellValue("TOTALE")
            val rowDD11 = sheet2.getRow(10)
            val cellDD11 = rowDD11.createCell(3)
            if (roa_ind_ros!!.getText()
                    .toString().length != 0 && ricavi_netti_vendita_ind_ros!!.getText()
                    .toString().length != 0
            ) {
                val redditogestioneoperativoros = roa_ind_ros!!.getText().toString().toDouble()
                val ricavinettivenditaros =
                    ricavi_netti_vendita_ind_ros!!.getText().toString().toDouble()
                cellDD9.setCellValue(redditogestioneoperativoros)
                cellDD10.setCellValue(ricavinettivenditaros)
                cellDD11.cellFormula = "D9/D10"
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(10, 10, 3, 3),
                    sheet2
                )
            } else {
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(10, 10, 3, 3),
                    sheet2
                )
            }
            val rowAA13 = sheet2.getRow(12)
            val cellAA13 = rowAA13.createCell(0)
            cellAA13.setCellStyle(style)
            cellAA13.setCellValue("ROE")
            val rowDD13 = sheet2.getRow(12)
            val cellDD13 = rowDD13.createCell(3)
            val rowCC13 = sheet2.getRow(12)
            val cellCC13 = rowCC13.createCell(2)
            cellCC13.setCellValue("REDDITO NETTO UTILE")
            val rowCC14 = sheet2.getRow(13)
            val cellCC14 = rowCC14.createCell(2)
            cellCC14.setCellValue("PATRIMONIO NETTO")
            val rowDD14 = sheet2.getRow(13)
            val cellDD14 = rowDD14.createCell(3)
            val rowAA15 = sheet2.getRow(14)
            val cellAA15 = rowAA15.createCell(0)
            cellAA15.setCellValue("TOTALE")
            val rowDD15 = sheet2.getRow(14)
            val cellDD15 = rowDD15.createCell(3)
            if (reddito_netto_utile_vendita_ind_roe!!.getText()
                    .toString().length != 0 && patrimonio_netto_ind_roe!!.getText()
                    .toString().length != 0
            ) {
                val redditonettoutileroe =
                    reddito_netto_utile_vendita_ind_roe!!.getText().toString().toDouble()
                val patrimonionettoroe = patrimonio_netto_ind_roe!!.getText().toString().toDouble()
                cellDD13.setCellValue(redditonettoutileroe)
                cellDD14.setCellValue(patrimonionettoroe)
                cellDD15.cellFormula = "D13/D14"
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(14, 14, 3, 3),
                    sheet2
                )
            } else {
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(14, 14, 3, 3),
                    sheet2
                )
            }
            val rowAA17 = sheet2.getRow(16)
            val cellAA17 = rowAA17.createCell(0)
            cellAA17.setCellStyle(style)
            cellAA17.setCellValue("TURNOVER")
            val rowDD17 = sheet2.getRow(16)
            val cellDD17 = rowDD17.createCell(3)
            val rowCC17 = sheet2.getRow(16)
            val cellCC17 = rowCC17.createCell(2)
            cellCC17.setCellValue("RICAVI NETTI DI VENDITA")
            val rowCC18 = sheet2.getRow(17)
            val cellCC18 = rowCC18.createCell(2)
            cellCC18.setCellValue("ATTIVO TOTALE NETTO")
            val rowDD18 = sheet2.getRow(17)
            val cellDD18 = rowDD18.createCell(3)
            val rowAA19 = sheet2.getRow(18)
            val cellAA19 = rowAA19.createCell(0)
            cellAA19.setCellValue("TOTALE")
            val rowDD19 = sheet2.getRow(18)
            val cellDD19 = rowDD19.createCell(3)
            if (ricavi_netti_vendita_ind_turnover!!.getText()
                    .toString().length != 0 && attivo_totale_netto_ind_turnover!!.getText()
                    .toString().length != 0
            ) {
                val ricavinettivenditaturnover =
                    ricavi_netti_vendita_ind_turnover!!.getText().toString().toDouble()
                val attivototalenettoturnover =
                    attivo_totale_netto_ind_turnover!!.getText().toString().toDouble()
                cellDD17.setCellValue(ricavinettivenditaturnover)
                cellDD18.setCellValue(attivototalenettoturnover)
                cellDD19.cellFormula = "D17/D18"
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(18, 18, 3, 3),
                    sheet2
                )
            } else {
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(18, 18, 3, 3),
                    sheet2
                )
            }

            // Bordo medio destro dopo aver inserito i valori
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(4, 5, 3, 3), sheet2)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(8, 9, 3, 3), sheet2)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(12, 13, 3, 3), sheet2)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(16, 17, 3, 3), sheet2)

            /*
    
    
                    FOGLIO 3
    
    
                     */
            val sheet3 = workbook.createSheet("INDICI DI LIQUIDITÀ")

            // Ciclo per creare le righe con i valori
            for (i in 0..10) {
                val row = sheet3.createRow(i)
                for (j in 0..3) {
                    val cell = row.createCell(j)
                }
            }

            // Imposta la larghezza delle colonne
            sheet3.setColumnWidth(0, 5000)
            sheet3.setColumnWidth(1, 5000)
            sheet3.setColumnWidth(2, 5000)
            sheet3.setColumnWidth(3, 5000)

            // Bordo piccolo
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet3)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet3)
            RegionUtil.setBorderLeft(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet3)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(0, 0, 0, 3), sheet3)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 1), sheet3)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet3)
            RegionUtil.setBorderRight(BorderStyle.THIN, CellRangeAddress(2, 2, 3, 3), sheet3)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(2, 2, 0, 3), sheet3)
            RegionUtil.setBorderBottom(BorderStyle.THIN, CellRangeAddress(1, 1, 0, 0), sheet3)
            RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress(3, 3, 0, 0), sheet3)


            //Bordo medio alto
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet3)
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, CellRangeAddress(11, 11, 0, 3), sheet3)

            //Bordo medio basso
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(3, 3, 0, 3), sheet3)
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, CellRangeAddress(7, 7, 0, 3), sheet3)

            // Unisci celle
            sheet3.addMergedRegion(CellRangeAddress(0, 0, 0, 3))
            sheet3.addMergedRegion(CellRangeAddress(2, 2, 0, 2))
            sheet3.addMergedRegion(CellRangeAddress(4, 4, 0, 1))
            sheet3.addMergedRegion(CellRangeAddress(8, 8, 0, 1))

            // Popolamento celle
            val rowAAA1 = sheet3.getRow(0)
            val cellAAA1 = rowAAA1.createCell(0)
            cellAAA1.setCellValue("INDICI DI BILANCIO - LIQUIDITÀ")
            cellAAA1.setCellStyle(horizontalstyle)
            val rowAAA3 = sheet3.getRow(2)
            val cellAAA3 = rowAAA3.createCell(0)
            cellAAA3.setCellValue("INDICI DI LIQUIDITÀ")
            cellAAA3.setCellStyle(horizontalstyle)
            val rowAAA5 = sheet3.getRow(4)
            val cellAAA5 = rowAAA5.createCell(0)
            cellAAA5.setCellValue("INDICE DI LIQUIDITÀ PRIMARIA")
            cellAAA5.setCellStyle(style)
            val rowDDD5 = sheet3.getRow(4)
            val cellDDD5 = rowDDD5.createCell(3)
            val rowCCC5 = sheet3.getRow(4)
            val cellCCC5 = rowCCC5.createCell(2)
            cellCCC5.setCellValue(
                """
    CASSA + BANCA + CREDITI
    VERSO CLIENTI
    """.trimIndent()
            )
            val rowCCC6 = sheet3.getRow(5)
            val cellCCC6 = rowCCC6.createCell(2)
            cellCCC6.setCellValue("PASSIVITÀ A BREVE TERMINE")
            val rowDDD6 = sheet3.getRow(5)
            val cellDDD6 = rowDDD6.createCell(3)
            val rowAAA7 = sheet3.getRow(6)
            val cellAAA7 = rowAAA7.createCell(0)
            cellAAA7.setCellValue("TOTALE")
            val rowDDD7 = sheet3.getRow(6)
            val cellDDD7 = rowDDD7.createCell(3)
            if (cassa_ind_liquidità_1!!.getText()
                    .toString().length != 0 || banca_ind_liquidità_1!!.getText()
                    .toString().length != 0 || crediti_ind_liquidità_1!!.getText()
                    .toString().length != 0 && passività_ind_liquidità_1!!.getText()
                    .toString().length != 0
            ) {
                val cassaliquidità = cassa_ind_liquidità_1!!.getText().toString().toDouble()
                val bancaliquidità = banca_ind_liquidità_1!!.getText().toString().toDouble()
                val creditiliquidità = crediti_ind_liquidità_1!!.getText().toString().toDouble()
                val passivitàabreveliquidità =
                    passività_ind_liquidità_1!!.getText().toString().toDouble()
                cellDDD5.setCellValue(cassaliquidità + bancaliquidità + creditiliquidità)
                cellDDD6.setCellValue(passivitàabreveliquidità)
                cellDDD7.cellFormula = "D5/D6"
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet3)
            } else {
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(6, 6, 3, 3), sheet3)
            }
            val rowAAA9 = sheet3.getRow(8)
            val cellAAA9 = rowAAA9.createCell(0)
            cellAAA9.setCellValue("INDICE DI LIQUIDITÀ SECONDARIA")
            cellAAA9.setCellStyle(style)
            val rowDDD9 = sheet3.getRow(8)
            val cellDDD9 = rowDDD9.createCell(3)
            val rowCCC9 = sheet3.getRow(8)
            val cellCCC9 = rowCCC9.createCell(2)
            cellCCC9.setCellValue("ATTIVO A BREVE")
            val rowCCC10 = sheet3.getRow(9)
            val cellCCC10 = rowCCC10.createCell(2)
            cellCCC10.setCellValue("PASSIVITÀ A BREVE TERMINE")
            val rowDDD10 = sheet3.getRow(9)
            val cellDDD10 = rowDDD10.createCell(3)
            val rowAAA11 = sheet3.getRow(10)
            val cellAAA11 = rowAAA11.createCell(0)
            cellAAA11.setCellValue("TOTALE")
            val rowDDD11 = sheet3.getRow(10)
            val cellDDD11 = rowDDD11.createCell(3)
            if (attività_ind_liquidità_2!!.getText()
                    .toString().length != 0 && passività_ind_liquidità_2!!.getText()
                    .toString().length != 0
            ) {
                val attivoabreveliquidità = attività_ind_liquidità_2!!.getText().toString().toDouble()
                val passivitàabreveliquidità2 =
                    passività_ind_liquidità_2!!.getText().toString().toDouble()
                cellDDD9.setCellValue(attivoabreveliquidità)
                cellDDD10.setCellValue(passivitàabreveliquidità2)
                cellDDD11.cellFormula = "D9/D10"
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(10, 10, 3, 3),
                    sheet3
                )
            } else {
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress(10, 10, 3, 3),
                    sheet3
                )
            }

            // Bordo medio destro dopo aver inserito i valori
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(4, 5, 3, 3), sheet3)
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, CellRangeAddress(8, 9, 3, 3), sheet3)

            // Esegui il salvataggio del file
            fileSaver.execute(workbook)
        }
    }
}
