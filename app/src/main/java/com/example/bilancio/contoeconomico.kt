package com.example.bilancio

import ContoEconomicoSaver
import android.Manifest
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.PopupMenu
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class contoeconomico() : AppCompatActivity() {
    var context: Context = this

    private fun showPopupMenu(v: View) {
        val popup = PopupMenu(this, v)
        popup.show()
    }

    var r5: EditText? = null
    var r6: EditText? = null
    var r7: EditText? = null
    var r8: EditText? = null
    var r9: EditText? = null
    var r10: EditText? = null
    var r11: EditText? = null
    var r12: EditText? = null
    var r13: EditText? = null
    var r14: EditText? = null
    var r16: EditText? = null
    var r17: EditText? = null
    var r18: EditText? = null
    var r20: EditText? = null
    var r21: EditText? = null
    var r22: EditText? = null
    var r23: EditText? = null
    var r24: EditText? = null
    var r25: EditText? = null
    var r26: EditText? = null
    var r27: EditText? = null
    var r32: EditText? = null
    var r33: EditText? = null
    var r34: EditText? = null
    var r35: EditText? = null
    var r36: EditText? = null
    var r37: EditText? = null
    var r38: EditText? = null
    var r41: EditText? = null
    var r42: EditText? = null
    var r43: EditText? = null
    var r44: EditText? = null
    var r46: EditText? = null
    var r47: EditText? = null
    var r48: EditText? = null
    var r49: EditText? = null
    var r55: EditText? = null
    var r56: EditText? = null
    var r59: EditText? = null
    override fun onCreate(savedInstanceState: Bundle?) {

        val fileSaver = ContoEconomicoSaver(this)
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_contoeconomico)


        val popup = findViewById<Button>(R.id.popup_button_ce)

        popup.setOnClickListener(object : View.OnClickListener {
            override fun onClick(view: View) {
                val popupMenu = PopupMenu(this@contoeconomico, view)
                popupMenu.menuInflater.inflate(R.menu.popup_menu_ce, popupMenu.menu)

                popupMenu.setOnMenuItemClickListener(PopupMenu.OnMenuItemClickListener { item ->
                    when (item.itemId) {
                        R.id.menu_item_conto_economico -> {
                            val intent = Intent(this@contoeconomico, statopatrimoniale::class.java)
                            startActivity(intent)
                            true
                        }

                        R.id.menu_item_home -> {
                            val intent1 = Intent(this@contoeconomico, MainActivity::class.java)
                            startActivity(intent1)
                            true
                        }

                        else -> false
                    }
                })

                popupMenu.show()
            }
        })
        if ((ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED)
        ) {
            ActivityCompat.requestPermissions(
                this, arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),
                REQUEST_CODE_WRITE_EXTERNAL_STORAGE
            )
        }
        val salvace = findViewById<Button>(R.id.salvace)
        salvace.setOnClickListener(object : View.OnClickListener {
            override fun onClick(v: View) {
                val fileSaver = ContoEconomicoSaver(context)
                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("Conto Economico")

                val style: CellStyle = workbook.createCellStyle()
                style.alignment = HorizontalAlignment.CENTER
                val font: Font = workbook.createFont()
                font.bold = true

                val grassetto: CellStyle = workbook.createCellStyle()
                grassetto.setFont(font)
                val alcentro: CellStyle = workbook.createCellStyle()
                alcentro.alignment = HorizontalAlignment.CENTER
                alcentro.setFont(font)
                val stile: CellStyle = workbook.createCellStyle()
                stile.alignment = HorizontalAlignment.CENTER
                stile.setFont(font)
                sheet.addMergedRegion(CellRangeAddress.valueOf("A1:C1"))
                val row1 = sheet.createRow(0)
                val A1 = row1.createCell(0)
                val B1 = row1.createCell(1)
                val C1 = row1.createCell(2)
                A1.setCellValue("CONTO ECONOMICO")
                A1.setCellStyle(alcentro)
                //2 RIGA
                val row2 = sheet.createRow(1)
                val A2 = row2.createCell(0)
                val B2 = row2.createCell(1)
                val C2 = row2.createCell(2)
                val row3 = sheet.createRow(2)
                val A3 = row3.createCell(0)
                val B3 = row3.createCell(1)
                val C3 = row3.createCell(2)
                A3.setCellValue("A)")
                A3.setCellStyle(grassetto)
                B3.setCellValue("VALORE DELLA PRODUZIONE")
                B3.setCellStyle(grassetto)
                C3.cellFormula = "C4+C5+C6+C7+C8"
                C3.setCellStyle(grassetto)
                val row4 = sheet.createRow(3)
                val A4 = row4.createCell(0)
                val B4 = row4.createCell(1)
                val C4 = row4.createCell(2)
                A4.setCellValue("1)")
                B4.setCellValue("Ricavi delle vendite e delle prestazioni")
                r5 = findViewById(R.id.edittext_ricavi)
                if (r5!!.text.toString().isNotEmpty()) {
                    ricavi = r5!!.getText().toString().toDouble()
                    C4.setCellValue(ricavi)
                }
                val row5 = sheet.createRow(4)
                val A5 = row5.createCell(0)
                val B5 = row5.createCell(1)
                val C5 = row5.createCell(2)
                A5.setCellValue("2")
                B5.setCellValue("Variazione rimanenze di prodotti in corso di lavorazione, semil. e finiti")
                r6 = findViewById(R.id.edittext_variazione_rimanenze)
                if (r6!!.text.toString().isNotEmpty()) {
                    val riga6 = r6!!.getText().toString().toDouble()
                    C5.setCellValue(riga6)
                }
                val row6 = sheet.createRow(5)
                val A6 = row6.createCell(0)
                val B6 = row6.createCell(1)
                val C6 = row6.createCell(2)
                A6.setCellValue("3)")
                B6.setCellValue("Variazione dei lavori in corso su ordinazione")
                r7 = findViewById(R.id.edittext_variazione_lavori_in_corso)
                if (r7!!.text.toString().length != 0) {
                    val riga7 = r7!!.getText().toString().toDouble()
                    C6.setCellValue(riga7)
                }
                val row7 = sheet.createRow(6)
                val A7 = row7.createCell(0)
                val B7 = row7.createCell(1)
                val C7 = row7.createCell(2)
                A7.setCellValue("4)")
                B7.setCellValue("Incrementi di immobilizzazioni per lavori interni")
                r8 = findViewById(R.id.edittext_incrementi_immobilizzazioni)
                if (r8!!.text.toString().length != 0) {
                    val riga8 = r8!!.getText().toString().toDouble()
                    C7.setCellValue(riga8)
                }
                val row8 = sheet.createRow(7)
                val A8 = row8.createCell(0)
                val B8 = row8.createCell(1)
                val C8 = row8.createCell(2)
                A8.setCellValue("5)")
                B8.setCellValue("Altri ricavi e proventi")
                r9 = findViewById(R.id.edittext_altri_ricavi)
                if (r9!!.text.toString().length != 0) {
                    val riga9 = r9!!.getText().toString().toDouble()
                    C8.setCellValue(riga9)
                }
                val row9 = sheet.createRow(8)
                val A9 = row9.createCell(0)
                val B9 = row9.createCell(1)
                val C9 = row9.createCell(2)
                A9.setCellValue("B)")
                A9.setCellStyle(grassetto)
                B9.setCellValue("COSTI DELLA PRODUZIONE")
                B9.setCellStyle(grassetto)
                C9.cellFormula = "C10+C11+C12+C13+C19+C25+C24+C26+C27"
                C9.setCellStyle(grassetto)
                //10 RIGA
                val row10 = sheet.createRow(9)
                val A10 = row10.createCell(0)
                val B10 = row10.createCell(1)
                val C10 = row10.createCell(2)
                A10.setCellValue("6)")
                B10.setCellValue("Costi per acquisti di materie prime, sussidiarie, di consumo e di merci")
                r10 = findViewById(R.id.edittext_costi_per_acquisti)
                if (r10!!.text.toString().length != 0) {
                    val riga10 = r10!!.getText().toString().toDouble()
                    C10.setCellValue(riga10)
                }

                //11 RIGA
                val row11 = sheet.createRow(10)
                val A11 = row11.createCell(0)
                val B11 = row11.createCell(1)
                val C11 = row11.createCell(2)
                A11.setCellValue("7)")
                B11.setCellValue("Costi per servizi")
                r11 = findViewById(R.id.edittext_costi_per_servizi)
                if (r11!!.text.toString().length != 0) {
                    val riga11 = r11!!.getText().toString().toDouble()
                    C11.setCellValue(riga11)
                }

//12 RIGA
                val row12 = sheet.createRow(11)
                val A12 = row12.createCell(0)
                val B12 = row12.createCell(1)
                val C12 = row12.createCell(2)
                A12.setCellValue("8)")
                B12.setCellValue("Costi per godimento di beni di terzi")
                r12 = findViewById(R.id.edittext_costi_per_godimento)
                if (r12!!.text.toString().length != 0) {
                    val riga12 = r12!!.getText().toString().toDouble()
                    C12.setCellValue(riga12)
                }

//13 RIGA
                val row13 = sheet.createRow(12)
                val A13 = row13.createCell(0)
                val B13 = row13.createCell(1)
                val C13 = row13.createCell(2)
                A13.setCellValue("9")
                B13.setCellValue("Costi per il personale")
                B13.setCellStyle(grassetto)
                C13.cellFormula = "C14+C15+C16+C17+C18"
                C13.setCellStyle(grassetto)

//14 RIGA
                val row14 = sheet.createRow(13)
                val A14 = row14.createCell(0)
                val B14 = row14.createCell(1)
                val C14 = row14.createCell(2)
                B14.setCellValue("a) Salari e stipendi")
                r13 = findViewById(R.id.edittext_salari)
                if (r13!!.text.toString().length != 0) {
                    val salari = r13!!.getText().toString().toDouble()
                    C14.setCellValue(salari)
                }

//15 RIGA
                val row15 = sheet.createRow(14)
                val A15 = row15.createCell(0)
                val B15 = row15.createCell(1)
                val C15 = row15.createCell(2)
                B15.setCellValue("b) Oneri sociali")
                r14 = findViewById(R.id.edittext_oneri_sociali)
                if (r14!!.text.toString().length != 0) {
                    val oneri = r14!!.getText().toString().toDouble()
                    C15.setCellValue(oneri)
                }

//16 RIGA
                val row16 = sheet.createRow(15)
                val A16 = row16.createCell(0)
                val B16 = row16.createCell(1)
                val C16 = row16.createCell(2)
                B16.setCellValue("c) Trattamento di fine rapporto")
                r16 = findViewById(R.id.edittext_trattamento_fine_rapporto)
                if (r16!!.text.toString().length != 0) {
                    val tfr = r16!!.getText().toString().toDouble()
                    C16.setCellValue(tfr)
                }

//17 RIGA
                val row17 = sheet.createRow(16)
                val A17 = row17.createCell(0)
                val B17 = row17.createCell(1)
                val C17 = row17.createCell(2)
                B17.setCellValue("d) Trattamento di quiescenza e simili")
                r17 = findViewById(R.id.edittext_trattamento_quiescenza)
                if (r17!!.text.toString().length != 0) {
                    val quiescenza = r17!!.getText().toString().toDouble()
                    C17.setCellValue(quiescenza)
                }
                //18 RIGA
                val row18 = sheet.createRow(17)
                val A18 = row18.createCell(0)
                val B18 = row18.createCell(1)
                val C18 = row18.createCell(2)
                B18.setCellValue("e) Altri costi")
                r18 = findViewById(R.id.edittext_altri_costi)
                if (r18!!.text.toString().length != 0) {
                    val altricosti = r18!!.getText().toString().toDouble()
                    C18.setCellValue(altricosti)
                }

//19 RIGA
                val row19 = sheet.createRow(18)
                val A19 = row19.createCell(0)
                val B19 = row19.createCell(1)
                val C19 = row19.createCell(2)
                A19.setCellValue("10)")
                B19.setCellValue("Ammortamenti e svalutazioni")
                B19.setCellStyle(grassetto)
                C19.cellFormula = "C20+C21+C22+C23"
                C19.setCellStyle(grassetto)

//20 RIGA
                val row20 = sheet.createRow(19)
                val A20 = row20.createCell(0)
                val B20 = row20.createCell(1)
                val C20 = row20.createCell(2)
                B20.setCellValue("a) Ammortamento immobilizzazioni immateriali")
                r20 = findViewById(R.id.edittext_ammortamento_immobilizzazioni_immateriali)
                if (r20!!.text.toString().length != 0) {
                    val ammortamento = r20!!.getText().toString().toDouble()
                    C20.setCellValue(ammortamento)
                }
                //21 RIGA
                val row21 = sheet.createRow(20)
                val A21 = row21.createCell(0)
                val B21 = row21.createCell(1)
                val C21 = row21.createCell(2)
                B21.setCellValue("b) Ammortamento immobilizzazioni materiali")
                r21 = findViewById(R.id.edittext_ammortamento_immobilizzazioni_materiali)
                if (r21!!.text.toString().length != 0) {
                    val ammortamentob = r21!!.getText().toString().toDouble()
                    C21.setCellValue(ammortamentob)
                }

//22 RIGA
                val row22 = sheet.createRow(21)
                val A22 = row22.createCell(0)
                val B22 = row22.createCell(1)
                val C22 = row22.createCell(2)
                B22.setCellValue("c) Altre svalutazioni delle immobilizzazioni")
                r22 = findViewById(R.id.edittext_altre_svalutazioni_immobilizzazioni)
                if (r22!!.text.toString().length != 0) {
                    val ammortamentoc = r22!!.getText().toString().toDouble()
                    C22.setCellValue(ammortamentoc)
                }


                //23 RIGA
                val row23 = sheet.createRow(22)
                val A23 = row23.createCell(0)
                val B23 = row23.createCell(1)
                val C23 = row23.createCell(2)
                B23.setCellValue("d) Svalutazione dei crediti compresi all'attivo circolante e delle disp. Liq.")
                r23 = findViewById(R.id.edittext_svalutazione_crediti)
                if (r23!!.text.toString().length != 0) {
                    val ammortamentod = r23!!.getText().toString().toDouble()
                    C23.setCellValue(ammortamentod)
                }

//24 RIGA
                val row24 = sheet.createRow(23)
                val A24 = row24.createCell(0)
                val B24 = row24.createCell(1)
                val C24 = row24.createCell(2)
                A24.setCellValue("11)")
                B24.setCellValue("Variazione delle rimanenze di materie prime, sussidiarie, di consumo e di merci")
                r24 = findViewById(R.id.edittext_variazione_rimanenzee)
                if (r24!!.text.toString().length != 0) {
                    val rimanenzeee = r24!!.getText().toString().toDouble()
                    C24.setCellValue(rimanenzeee)
                }
                //25 RIGA
                val row25 = sheet.createRow(24)
                val A25 = row25.createCell(0)
                val B25 = row25.createCell(1)
                val C25 = row25.createCell(2)
                A25.setCellValue("12)")
                B25.setCellValue("Accantonamenti per rischi")
                r25 = findViewById(R.id.edittext_accantonamenti_rischi)
                if (r25!!.text.toString().length != 0) {
                    val rischi = r25!!.getText().toString().toDouble()
                    C25.setCellValue(rischi)
                }
                //26 RIGA
                val row26 = sheet.createRow(25)
                val A26 = row26.createCell(0)
                val B26 = row26.createCell(1)
                val C26 = row26.createCell(2)
                A26.setCellValue("13)")
                B26.setCellValue("Altri accantonamenti")
                r26 = findViewById(R.id.edittext_altri_accantonamenti)
                if (r26!!.text.toString().length != 0) {
                    val altririschi = r26!!.getText().toString().toDouble()
                    C26.setCellValue(altririschi)
                }
                //27 RIGA
                val row27 = sheet.createRow(26)
                val A27 = row27.createCell(0)
                val B27 = row27.createCell(1)
                val C27 = row27.createCell(2)
                A27.setCellValue("14)")
                B27.setCellValue("Oneri diversi di gestione")
                r27 = findViewById(R.id.edittext_oneri_diversi_gestione)
                if (r27!!.text.toString().length != 0) {
                    val oneridiversi = r27!!.getText().toString().toDouble()
                    C27.setCellValue(oneridiversi)
                }


//28 RIGA
                val row28 = sheet.createRow(27)
                val A28 = row28.createCell(0)
                val B28 = row28.createCell(1)
                val C28 = row28.createCell(2)
                B28.setCellValue("Differenza tra valore e costi della produzione (A-B)")
                B28.setCellStyle(grassetto)
                C28.cellFormula = "C3-C9"
                C28.setCellStyle(grassetto)


//29 RIGA
                val row29 = sheet.createRow(28)
                val A29 = row29.createCell(0)
                val B29 = row29.createCell(1)
                val C29 = row29.createCell(2)

//30 RIGA
                val row30 = sheet.createRow(29)
                val A30 = row30.createCell(0)
                val B30 = row30.createCell(1)
                val C30 = row30.createCell(2)
                A30.setCellValue("C)")
                A30.setCellStyle(grassetto)
                B30.setCellValue("PROVENTI E ONERI FINANZIARI")
                B30.setCellStyle(grassetto)
                C30.cellFormula = "C31+C37+C38"
                C30.setCellStyle(grassetto)

//31 RIGA
                val row31 = sheet.createRow(30)
                val A31 = row31.createCell(0)
                val B31 = row31.createCell(1)
                val C31 = row31.createCell(2)
                A31.setCellValue("15)")
                B31.setCellValue("Proventi delle partecipazioni")
                B31.setCellStyle(grassetto)
                C31.cellFormula = "C32+C33+C34+C35+C36"
                C31.setCellStyle(grassetto)
                //32 RIGA
                val row32 = sheet.createRow(31)
                val A32 = row32.createCell(0)
                val B32 = row32.createCell(1)
                val C32 = row32.createCell(2)
                B32.setCellValue("a) in imprese controllate")
                r32 = findViewById(R.id.edittext_imp_controllate)
                if (r32!!.text.toString().length != 0) {
                    val riga32 = r32!!.getText().toString().toDouble()
                    C32.setCellValue(riga32)
                }


//33 RIGA
                val row33 = sheet.createRow(32)
                val A33 = row33.createCell(0)
                val B33 = row33.createCell(1)
                val C33 = row33.createCell(2)
                B33.setCellValue("b) in imprese collegate")
                r33 = findViewById(R.id.edittext_imp_collegate)
                if (r33!!.text.toString().length != 0) {
                    val riga33 = r33!!.getText().toString().toDouble()
                    C33.setCellValue(riga33)
                }


                // Creating row 34
                val row34 = sheet.createRow(33)
                val A34 = row34.createCell(0)
                val B34 = row34.createCell(1)
                val C34 = row34.createCell(2)
                B34.setCellValue("c) in imprese controllanti")
                r34 = findViewById(R.id.edittext_imp_controllanti)
                if (r34!!.text.toString().length != 0) {
                    val riga34 = r34!!.getText().toString().toDouble()
                    C34.setCellValue(riga34)
                }


// Creating row 35
                val row35 = sheet.createRow(34)
                val A35 = row35.createCell(0)
                val B35 = row35.createCell(1)
                val C35 = row35.createCell(2)
                B35.setCellValue("d) in imprese sottoposte al controllo delle controllanti")
                r35 = findViewById(R.id.edittext_imp_sottoposte_controllo)
                if (r35!!.text.toString().length != 0) {
                    val riga35 = r35!!.getText().toString().toDouble()
                    C35.setCellValue(riga35)
                }

// Creating row 36
                val row36 = sheet.createRow(35)
                val A36 = row36.createCell(0)
                val B36 = row36.createCell(1)
                val C36 = row36.createCell(2)
                B36.setCellValue("e) in altre imprese")
                r36 = findViewById(R.id.edittext_altre_imprese)
                if (r36!!.text.toString().length != 0) {
                    val riga36 = r36!!.getText().toString().toDouble()
                    C36.setCellValue(riga36)
                }

// Creating row 37
                val row37 = sheet.createRow(36)
                val A37 = row37.createCell(0)
                val B37 = row37.createCell(1)
                val C37 = row37.createCell(2)
                A37.setCellValue("16)")
                B37.setCellValue("Altri proventi finanziari")
                r37 = findViewById(R.id.edittext_altri_proventi_finanziari)
                if (r37!!.text.toString().length != 0) {
                    val riga37 = r37!!.getText().toString().toDouble()
                    C37.setCellValue(riga37)
                }

// Creating row 38
                val row38 = sheet.createRow(37)
                val A38 = row38.createCell(0)
                val B38 = row38.createCell(1)
                val C38 = row38.createCell(2)
                A38.setCellValue("17)")
                B38.setCellValue("Interessi ed altri oneri finanziari")
                r38 = findViewById(R.id.edittext_interessi_oneri_finanziari)
                if (r38!!.text.toString().length != 0) {
                    val riga38 = r38!!.getText().toString().toDouble()
                    C38.setCellValue(riga38)
                }

// Creating row 39
                val row39 = sheet.createRow(38)
                val A39 = row39.createCell(0)
                val B39 = row39.createCell(1)
                val C39 = row39.createCell(2)
                A39.setCellValue("D)")
                A39.setCellStyle(grassetto)
                B39.setCellValue("RETTIFICHE DI VALORE DI ATTIVITA' FINANZIARIE")
                B39.setCellStyle(grassetto)

// Creating row 40
                val row40 = sheet.createRow(39)
                val A40 = row40.createCell(0)
                val B40 = row40.createCell(1)
                val C40 = row40.createCell(2)
                A40.setCellValue("18)")
                B40.setCellValue("rivalutazione di attività finanziarie")
                C40.cellFormula = "C41+C42+C43+C44"
                B40.setCellStyle(grassetto)
                C40.setCellStyle(grassetto)
                if (C40.numericCellValue == 0.0) {
                    C40.setCellValue("1")
                }

// Creating row 41
                val row41 = sheet.createRow(40)
                val A41 = row41.createCell(0)
                val B41 = row41.createCell(1)
                val C41 = row41.createCell(2)
                B41.setCellValue("a) di partecipazioni ")
                r41 = findViewById(R.id.edittext_partecipazioni)
                if (r41!!.text.toString().length != 0) {
                    val riga41 = r41!!.getText().toString().toDouble()
                    C41.setCellValue(riga41)
                }

// Creating row 42
                val row42 = sheet.createRow(41)
                val A42 = row42.createCell(0)
                val B42 = row42.createCell(1)
                val C42 = row42.createCell(2)
                B42.setCellValue("b) di immobilizzazioni finanziarie che non sono partecipazioni")
                r42 = findViewById(R.id.edittext_immobili_finanziari)
                if (r42!!.text.toString().length != 0) {
                    val riga42 = r42!!.getText().toString().toDouble()
                    C42.setCellValue(riga42)
                }


// Creating row 43
                val row43 = sheet.createRow(42)
                val A43 = row43.createCell(0)
                val B43 = row43.createCell(1)
                val C43 = row43.createCell(2)
                B43.setCellValue("c) di titoli iscritti nell'attivo circolante che non sono partecipazioni")
                r43 = findViewById(R.id.edittext_titoli_attivo_circolante)
                if (r43!!.text.toString().length != 0) {
                    val riga43 = r43!!.getText().toString().toDouble()
                    C43.setCellValue(riga43)
                }

// Creating row 44
                val row44 = sheet.createRow(43)
                val A44 = row44.createCell(0)
                val B44 = row44.createCell(1)
                val C44 = row44.createCell(2)
                B44.setCellValue("d) di strumenti finanziari derivati")
                r44 = findViewById(R.id.edittext_strumenti_finanziari_derivati)
                if (r44!!.text.toString().length != 0) {
                    val riga44 = r44!!.getText().toString().toDouble()
                    C44.setCellValue(riga44)
                }

// Creating row 45
                val row45 = sheet.createRow(44)
                val A45 = row45.createCell(0)
                val B45 = row45.createCell(1)
                val C45 = row45.createCell(2)
                A45.setCellValue("19)")
                B45.setCellValue("svalutazioni di attività finanziarie")
                B45.setCellStyle(grassetto)
                C45.cellFormula = "C46+C47+C48+C49"
                C45.setCellStyle(grassetto)


// Creating row 46
                val row46 = sheet.createRow(45)
                val A46 = row46.createCell(0)
                val B46 = row46.createCell(1)
                val C46 = row46.createCell(2)
                B46.setCellValue("a) di partecipazioni ")
                r46 = findViewById(R.id.edittext_partecipazioni_2)
                if (r46!!.text.toString().length != 0) {
                    val riga46 = r46!!.getText().toString().toDouble()
                    C46.setCellValue(riga46)
                }


// Creating row 47
                val row47 = sheet.createRow(46)
                val A47 = row47.createCell(0)
                val B47 = row47.createCell(1)
                val C47 = row47.createCell(2)
                B47.setCellValue("b) di immobilizzazioni finanziarie che non sono partecipazioni")
                r47 = findViewById(R.id.edittext_immobili_finanziari_2)
                if (r47!!.text.toString().length != 0) {
                    val riga47 = r47!!.getText().toString().toDouble()
                    C47.setCellValue(riga47)
                }

// Creating row 48
                val row48 = sheet.createRow(47)
                val A48 = row48.createCell(0)
                val B48 = row48.createCell(1)
                val C48 = row48.createCell(2)
                B48.setCellValue("c) di titoli iscritti nell'attivo circolante che non sono partecipazioni")
                r48 = findViewById(R.id.edittext_titoli_attivo_circolante_2)
                if (r48!!.text.toString().length != 0) {
                    val riga48 = r48!!.getText().toString().toDouble()
                    C48.setCellValue(riga48)
                }

// Creating row 49
                val row49 = sheet.createRow(48)
                val A49 = row49.createCell(0)
                val B49 = row49.createCell(1)
                val C49 = row49.createCell(2)
                B49.setCellValue("d) di strumenti finanziari derivati")
                r49 = findViewById(R.id.edittext_strumenti_finanziari_derivati_2)
                if (r49!!.text.toString().length != 0) {
                    val riga49 = r49!!.getText().toString().toDouble()
                    C49.setCellValue(riga49)
                }

// Creating row 50
                val row50 = sheet.createRow(49)
                val A50 = row50.createCell(0)
                val B50 = row50.createCell(1)
                val C50 = row50.createCell(2)

// Creating row 51
                val row51 = sheet.createRow(50)
                val A51 = row51.createCell(0)
                val B51 = row51.createCell(1)
                val C51 = row51.createCell(2)
                B51.setCellValue("Totale delle rettifiche (18-19)")
                B51.setCellStyle(grassetto)
                C51.cellFormula = "C40-C45"
                C51.setCellStyle(grassetto)


// Creating row 52
                val row52 = sheet.createRow(51)
                val A52 = row52.createCell(0)
                val B52 = row52.createCell(1)
                val C52 = row52.createCell(2)

// Creating row 53
                val row53 = sheet.createRow(52)
                val A53 = row53.createCell(0)
                val B53 = row53.createCell(1)
                val C53 = row53.createCell(2)
                A53.setCellValue("E)")
                A53.setCellStyle(grassetto)
                B53.setCellValue("PROVENTI E ONERI STRAORDINARI")
                B53.setCellStyle(grassetto)
                C53.cellFormula = "C55+C56"
                C53.setCellStyle(grassetto)

// Creating row 54
                val row54 = sheet.createRow(53)
                val A54 = row54.createCell(0)
                val B54 = row54.createCell(1)
                val C54 = row54.createCell(2)

// Creating row 55
                val row55 = sheet.createRow(54)
                val A55 = row55.createCell(0)
                val B55 = row55.createCell(1)
                val C55 = row55.createCell(2)
                A55.setCellValue("20)")
                B55.setCellValue("Proventi straordinari")
                r55 = findViewById(R.id.proventi_straordinari)
                if (r55!!.text.toString().length != 0) {
                    val riga55 = r55!!.getText().toString().toDouble()
                    C55.setCellValue(riga55)
                }

// Creating row 56
                val row56 = sheet.createRow(55)
                val A56 = row56.createCell(0)
                val B56 = row56.createCell(1)
                val C56 = row56.createCell(2)
                A56.setCellValue("21)")
                B56.setCellValue("Oneri straordinari")
                r56 = findViewById(R.id.oneri_straordinari)
                if (r56!!.text.toString().length != 0) {
                    val riga56 = r56!!.getText().toString().toDouble()
                    C56.setCellValue(riga56)
                }

// Creating row 57
                val row57 = sheet.createRow(56)
                val A57 = row57.createCell(0)
                val B57 = row57.createCell(1)
                val C57 = row57.createCell(2)

// Creating row 58
                val row58 = sheet.createRow(57)
                val A58 = row58.createCell(0)
                val B58 = row58.createCell(1)
                val C58 = row58.createCell(2)
                B58.setCellValue("Risultato prima delle imposte")
                B58.setCellStyle(grassetto)
                C58.cellFormula = "C28+C30+C53"
                C58.setCellStyle(grassetto)


// Creating row 59
                val row59 = sheet.createRow(58)
                val A59 = row59.createCell(0)
                val B59 = row59.createCell(1)
                val C59 = row59.createCell(2)
                A59.setCellValue("22)")
                B59.setCellValue("Imposte sul reddito dell'esercizio")
                r59 = findViewById(R.id.edittext_imposte_reddito_esercizio)
                if (r59!!.text.toString().length != 0) {
                    val riga59 = r59!!.getText().toString().toDouble()
                    C59.setCellValue(riga59)
                }

// Creating row 60
                val row60 = sheet.createRow(59)
                val A60 = row60.createCell(0)
                val B60 = row60.createCell(1)
                val C60 = row60.createCell(2)
                A60.setCellValue("26)")
                A60.setCellStyle(grassetto)
                B60.setCellValue("Utile (perdita) dell'esercizio")
                B60.setCellStyle(grassetto)
                C60.cellFormula = "C58-C59"
                C60.setCellStyle(grassetto)

// Creating row 61
                val row61 = sheet.createRow(60)
                val A61 = row61.createCell(0)
                val B61 = row61.createCell(1)
                val C61 = row61.createCell(2)
                val evaluator: FormulaEvaluator = workbook.creationHelper.createFormulaEvaluator()
                evaluator.evaluateAll()
                RegionUtil.setBorderBottom(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A60:C60"), sheet
                )
                RegionUtil.setBorderTop(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A3:C3"), sheet
                )
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A3:A60"), sheet
                )
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("B3:B60"), sheet
                )
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("C3:C60"), sheet
                )
                RegionUtil.setBorderLeft(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A3:A60"), sheet
                )
                for (i in 4..60) {
                    RegionUtil.setBorderTop(
                        BorderStyle.THIN, CellRangeAddress.valueOf(
                            "A$i:C$i"
                        ), sheet
                    )
                }
                sheet.setColumnWidth(1, 20000)
                // Salvataggio del file
                fileSaver.execute(workbook)
            }
        })
    }

    companion object {
        var ricavi = 0.0
        private val REQUEST_CODE_WRITE_EXTERNAL_STORAGE = 1
    }
}