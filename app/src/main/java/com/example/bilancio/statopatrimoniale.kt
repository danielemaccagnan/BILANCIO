package com.example.bilancio

import android.Manifest
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.text.Editable
import android.text.TextWatcher
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.PopupMenu
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.example.bilancio.contoeconomico
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

class statopatrimoniale() : AppCompatActivity() {
    var context: Context = this
    private val fileSaverTask: StatoPatrimonialeSaver? = null

    private fun showPopupMenu(v: View) {
        val popup = PopupMenu(this, v)
        popup.show()
    }

    var creditiversosoci1: EditText? = null
    var immobilizzazioni1: EditText? = null
    var immobilizzazioni2: EditText? = null
    var immobilizzazioni3: EditText? = null
    var immobilizzazioni4: EditText? = null
    var immobilizzazioni5: EditText? = null
    var immobilizzazioni6: EditText? = null
    var immobilizzazioni7: EditText? = null
    var immobilizzazionimateriali1: EditText? = null
    var immobilizzazionimateriali2: EditText? = null
    var immobilizzazionimateriali3: EditText? = null
    var immobilizzazionimaterialifa: EditText? = null
    var immobilizzazionimateriali4: EditText? = null
    var immobilizzazionimateriali5: EditText? = null
    var immobilizzazionifinanziarie1: EditText? = null
    var immobilizzazionifinanziarie2: EditText? = null
    var immobilizzazionifinanziarie3: EditText? = null
    var immobilizzazionifinanziarie4: EditText? = null
    var rimanenze1: EditText? = null
    var rimanenze2: EditText? = null
    var rimanenze3: EditText? = null
    var rimanenze4: EditText? = null
    var rimanenze5: EditText? = null
    var cambialiattive: EditText? = null
    var crediti1punto1: EditText? = null
    var crediti1punto2: EditText? = null
    var crediti2: EditText? = null
    var crediti3: EditText? = null
    var crediti4: EditText? = null
    var crediti5: EditText? = null
    var crediti5punto2: EditText? = null
    var crediti5punto3: EditText? = null
    var crediti5punto4: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate1: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate2: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate3: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate3punto2: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate4: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate5: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate6: EditText? = null
    var immobilizzazionifinanziarienonimmobilizzate7: EditText? = null
    var disponibilitàliquide1: EditText? = null
    var disponibilitàliquide2: EditText? = null
    var disponibilitàliquide3: EditText? = null
    var rateieriscontiattivi1: EditText? = null
    var fatturedaemettere: EditText? = null
    var mutuipassivi: EditText? = null
    var mutuiattivi: EditText? = null
    var fornitoriacconti: EditText? = null
    var clientiacconti: EditText? = null
    var azionistisott: EditText? = null
    var erariociv: EditText? = null
    var fornitoriimmobilizzazioniimmaterialicacconti: EditText? = null
    var getFornitoriimmobilizzazionimaterialicacconti: EditText? = null
    var imballaggidurevoli: EditText? = null
    var crediticommdiversi: EditText? = null
    var clienticostianticipati: EditText? = null
    var cambialiallosconto: EditText? = null
    var cambialiallincasso: EditText? = null
    var cambialiinsolute: EditText? = null
    var fondorischisucrediti: EditText? = null
    var fondosvalutazionecrediti: EditText? = null
    var creditiperimposte: EditText? = null
    var creditiversoistitutiprevidenziali: EditText? = null
    var creditipercauzioni: EditText? = null
    var creditidiversi: EditText? = null
    var bancaccattivi: EditText? = null
    var bancaccpostali: EditText? = null
    var fondoresponsabilitàcivile: EditText? = null
    var fondomanutenzioniprogrammate: EditText? = null
    var debitipertfr: EditText? = null
    var bancacriba: EditText? = null
    var bancaccpassivi: EditText? = null
    var bancacanticipasufattura: EditText? = null
    var devitivaltrifinanziatori: EditText? = null
    var debiticommercialidiversi: EditText? = null
    var debitidaliquidare: EditText? = null
    var clienticacconti: EditText? = null
    var cambialipassive: EditText? = null
    var debitiperritenutedaversare: EditText? = null
    var debitiperiva: EditText? = null
    var debitiperimposte: EditText? = null
    var erariociva: EditText? = null
    var debitipercauzioni: EditText? = null
    var personalecretribuzione: EditText? = null
    var personalecliquidazione: EditText? = null
    var debitivfpensione: EditText? = null
    var cedenteccessione: EditText? = null
    var debitidiversi: EditText? = null
    var patrimonionetto1: EditText? = null
    var patrimonionetto2: EditText? = null
    var patrimonionetto3: EditText? = null
    var patrimonionetto4: EditText? = null
    var patrimonionetto5: EditText? = null
    var patrimonionetto6: EditText? = null
    var patrimonionetto7: EditText? = null
    var patrimonionetto8: EditText? = null
    var patrimonionetto9: EditText? = null
    var patrimonionetto10: EditText? = null
    var fondiperrischieoneri1: EditText? = null
    var fondiperrischieoneri2: EditText? = null
    var fondiperrischieoneri3: EditText? = null
    var fondiperrischieoneri4: EditText? = null
    var tfr1: EditText? = null
    var fatturedaricevere: EditText? = null
    var debiti1: EditText? = null
    var debiti2: EditText? = null
    var debiti3: EditText? = null
    var debiti4: EditText? = null
    var debiti4punto1: EditText? = null
    var debiti4punto2: EditText? = null
    var debiti5: EditText? = null
    var debiti5punto1: EditText? = null
    var debiti5punto2: EditText? = null
    var debiti6: EditText? = null
    var debiti7: EditText? = null
    var debiti8: EditText? = null
    var debiti9: EditText? = null
    var debiti10: EditText? = null
    var debiti11: EditText? = null
    var debiti11punto2: EditText? = null
    var debiti12: EditText? = null
    var debiti13: EditText? = null
    var debiti14: EditText? = null
    var rateieriscontipassivi1: EditText? = null
    override fun onCreate(savedInstanceState: Bundle?) {
        val context: Context = this


        val fileSaver = StatoPatrimonialeSaver(this)
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_statopatrimoniale)



        val popup = findViewById<Button>(R.id.popup_button)


        popup.setOnClickListener(object : View.OnClickListener {
            override fun onClick(view: View) {
                val popupMenu = PopupMenu(this@statopatrimoniale, view)
                popupMenu.menuInflater.inflate(R.menu.popup_menu, popupMenu.menu)

                popupMenu.setOnMenuItemClickListener(PopupMenu.OnMenuItemClickListener { item ->
                    when (item.itemId) {
                        R.id.menu_item_conto_economico -> {
                            val intent = Intent(this@statopatrimoniale, contoeconomico::class.java)
                            startActivity(intent)
                            true
                        }

                        R.id.menu_item_home -> {
                            val intent1 = Intent(this@statopatrimoniale, MainActivity::class.java)
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


        val debiti11punto2 = findViewById<EditText>(R.id.debiti_11_2)
        debiti11punto2.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    debiti11punto2.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    debiti11punto2.setSelection(debiti11punto2.text.length)
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    debiti11punto2.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    debiti11punto2.setSelection(debiti11punto2.text.length)
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })
        val patrimonionetto7 = findViewById<EditText>(R.id.patrimonionetto_7)
        patrimonionetto7.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    patrimonionetto7.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    patrimonionetto7.setSelection(patrimonionetto7.text.length)
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    patrimonionetto7.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    patrimonionetto7.setSelection(patrimonionetto7.text.length)
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })
        val patrimonionetto10 = findViewById<EditText>(R.id.patrimonionetto_10)
        patrimonionetto10.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    patrimonionetto10.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    patrimonionetto10.setSelection(patrimonionetto10.text.length)
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    patrimonionetto10.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    patrimonionetto10.setSelection(patrimonionetto10.text.length)
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })
        val immobilizzazionifinanziarienonimmobilizzate3punto2 =
            findViewById<EditText>(R.id.immobilizzazionifinanziarienonimmobilizzate_3_2)
        immobilizzazionifinanziarienonimmobilizzate3punto2.addTextChangedListener(object :
            TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setSelection(
                        immobilizzazionifinanziarienonimmobilizzate3punto2.text.length
                    )
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setSelection(
                        immobilizzazionifinanziarienonimmobilizzate3punto2.text.length
                    )
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })
        val crediti5 = findViewById<EditText>(R.id.crediti_5)
        crediti5.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    crediti5.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    crediti5.setSelection(crediti5.text.length)
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    crediti5.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    crediti5.setSelection(crediti5.text.length)
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })

        val debiti13 = findViewById<EditText>(R.id.debiti_13)
        debiti13.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {}
            override fun onTextChanged(charSequence: CharSequence, i: Int, i1: Int, i2: Int) {
                if (charSequence.toString().contains(".") && charSequence.toString()
                        .lastIndexOf(".") != charSequence.toString().indexOf(".")
                ) {
                    debiti13.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf(".")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf(".") + 1)
                    )
                    debiti13.setSelection(debiti13.text.length)
                }
                if (charSequence.toString().contains("-") && charSequence.toString()
                        .lastIndexOf("-") != charSequence.toString().indexOf("-")
                ) {
                    debiti13.setText(
                        charSequence.toString().substring(
                            0,
                            charSequence.toString().lastIndexOf("-")
                        ) + charSequence.toString()
                            .substring(charSequence.toString().lastIndexOf("-") + 1)
                    )
                    debiti13.setSelection(debiti13.text.length)
                }
            }

            override fun afterTextChanged(editable: Editable) {}
        })
        val salvaBilancio = findViewById<Button>(R.id.salva)
        salvaBilancio.setOnClickListener(object : View.OnClickListener {
            override fun onClick(v: View) {

                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("Stato Patrimoniale")
                var fileName: String = "statopatrimoniale.xlsx"
                var counter = 1
                var file = File(context.getExternalFilesDir(null), fileName)
                while (file.exists()) {
                    fileName = "statopatrimoniale($counter).xlsx"
                    file = File(context.getExternalFilesDir(null), fileName)
                    counter++
                }

                val fileSaver = StatoPatrimonialeSaver(context)
                val row = sheet.createRow(0)
                sheet.addMergedRegion(CellRangeAddress.valueOf("A1:F1"))

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
                val sottile: CellStyle = workbook.createCellStyle()
                sottile.borderRight = BorderStyle.THIN
                sottile.borderTop = BorderStyle.THIN


                val statopatrimoniale: Cell = row.createCell(0)
                statopatrimoniale.setCellValue("STATO PATRIMONIALE")
                statopatrimoniale.cellStyle = alcentro
                // 2 RIGA
                val row2 = sheet.createRow(1)
                val A2 = row2.createCell(0)
                val B2 = row2.createCell(1)
                val C2 = row2.createCell(2)
                val D2 = row2.createCell(3)
                val E2 = row2.createCell(4)
                val F2 = row2.createCell(5)
                sheet.addMergedRegion(CellRangeAddress.valueOf("A3:B3"))
                sheet.addMergedRegion(CellRangeAddress.valueOf("D3:E3"))
                sheet.addMergedRegion(CellRangeAddress.valueOf("A61:B61"))
                sheet.addMergedRegion(CellRangeAddress.valueOf("D61:E61"))
                // 3 RIGA
                val row3 = sheet.createRow(2)
                val cellA3 = row3.createCell(0)
                val cellB3 = row3.createCell(1)
                val cellC3 = row3.createCell(2)
                val cellD3 = row3.createCell(3)
                val cellE3 = row3.createCell(4)
                val cellF3 = row3.createCell(5)
                cellA3.setCellValue("ATTIVO")
                cellA3.setCellStyle(alcentro)
                cellC3.setCellValue("data")
                cellC3.setCellStyle(grassetto)
                cellD3.setCellValue("PASSIVO")
                cellD3.setCellStyle(alcentro)
                cellF3.setCellValue("data")
                cellF3.setCellStyle(grassetto)


                //5 RIGA
                val row5 = sheet.createRow(4)
                val cellA5 = row5.createCell(0)
                val cellB5 = row5.createCell(1)
                val cellC5 = row5.createCell(2)
                val cellD5 = row5.createCell(3)
                val cellE5 = row5.createCell(4)
                val cellF5 = row5.createCell(5)
                cellA5.setCellValue("A)")
                cellA5.setCellStyle(grassetto)
                cellB5.setCellValue("CREDITI VERSO SOCI PER VERSAMENTI ANCORA DOVUTI")
                cellB5.setCellStyle(grassetto)
                cellC5.setCellStyle(grassetto)
                cellD5.setCellValue("A)")
                cellD5.setCellStyle(grassetto)
                cellE5.setCellValue("PATRIMONIO NETTO")
                cellE5.setCellStyle(grassetto)
                cellF5.cellFormula = "SUM(F6:F16)"
                cellF5.setCellStyle(grassetto)
                creditiversosoci1 = findViewById(R.id.creditiversosoci_1)
                azionistisott = findViewById(R.id.azionisticsott)
                if (creditiversosoci1!!.getText().toString().length != 0 || azionistisott!!.getText()
                        .toString().length != 0
                ) {
                    val crevssoci1: Double =
                        if (creditiversosoci1!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else creditiversosoci1!!.getText().toString().toDouble()
                    val azionistisotto: Double =
                        if (azionistisott!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else azionistisott!!.getText().toString().toDouble()
                    cellC5.setCellValue(crevssoci1 + azionistisotto)
                }


                // SESTA RIGA
                val row6 = sheet.createRow(5)
                val cellE6 = row6.createCell(4)
                val cellF6 = row6.createCell(5)
                cellE6.setCellValue("I. Capitale sociale")
                patrimonionetto1 = findViewById(R.id.patrimonionetto_1)


                // SETTIMA RIGA
                val row7 = sheet.createRow(6)
                val cellA7 = row7.createCell(0)
                val cellB7 = row7.createCell(1)
                val cellC7 = row7.createCell(2)
                val cellE7 = row7.createCell(4)
                val cellF7 = row7.createCell(5)
                patrimonionetto2 = findViewById(R.id.patrimonionetto_2)
                if (patrimonionetto2!!.getText().toString().length != 0) {
                    val patnetto2 = patrimonionetto2!!.getText().toString().toDouble()
                    cellF7.setCellValue(patnetto2)
                }
                cellA7.setCellValue("B)")
                cellA7.setCellStyle(grassetto)
                cellB7.setCellValue("IMMOBILIZZAZIONI")
                cellB7.setCellStyle(grassetto)
                cellC7.cellFormula = "C8+C16+C23"
                cellC7.setCellStyle(grassetto)
                if (cellC7.numericCellValue == 0.0) {
                    cellC7.setCellValue("1")
                }
                cellE7.setCellValue("II. Riserva sovrapprezzo delle azioni")
                // 8 RIGA
                val row8 = sheet.createRow(7)
                val A8 = row8.createCell(0)
                val B8 = row8.createCell(1)
                val C8 = row8.createCell(2)
                val D8 = row8.createCell(3)
                val E8 = row8.createCell(4)
                val F8 = row8.createCell(5)
                patrimonionetto3 = findViewById(R.id.patrimonionetto_3)
                if (patrimonionetto3!!.getText().toString().length != 0) {
                    val patnetto3 = patrimonionetto3!!.getText().toString().toDouble()
                    F8.setCellValue(patnetto3)
                }
                A8.setCellValue("I)")
                A8.setCellStyle(grassetto)
                B8.setCellValue("Immobilizzazioni Immateriali")
                B8.setCellStyle(grassetto)
                C8.cellFormula = "SUM(C9:C15)"
                C8.setCellStyle(grassetto)
                E8.setCellValue("III. Riserva di rivalutazione")

                //9 RIGA
                val row9 = sheet.createRow(8)
                val A9 = row9.createCell(0)
                val B9 = row9.createCell(1)
                val C9 = row9.createCell(2)
                val D9 = row9.createCell(3)
                val E9 = row9.createCell(4)
                val F9 = row9.createCell(5)
                B9.setCellValue("1) Costi di impianto e ampliamento")
                E9.setCellValue("IV. Riserva legale")
                patrimonionetto4 = findViewById(R.id.patrimonionetto_4)
                if (patrimonionetto4!!.getText().toString().length != 0) {
                    val patnetto4 = patrimonionetto4!!.getText().toString().toDouble()
                    F9.setCellValue(patnetto4)
                }
                immobilizzazioni1 = findViewById(R.id.immobilizzazioni_1)
                if (immobilizzazioni1!!.getText().toString().length != 0) {
                    val imm1 = immobilizzazioni1!!.getText().toString().toDouble()
                    C9.setCellValue(imm1)
                }


                // 10 RIGA
                val row10 = sheet.createRow(9)
                val A10 = row10.createCell(0)
                val B10 = row10.createCell(1)
                val C10 = row10.createCell(2)
                val D10 = row10.createCell(3)
                val E10 = row10.createCell(4)
                val F10 = row10.createCell(5)
                immobilizzazioni2 = findViewById(R.id.immobilizzazioni_2)
                if (immobilizzazioni2!!.getText().toString().length != 0) {
                    val imm2 = immobilizzazioni2!!.getText().toString().toDouble()
                    C10.setCellValue(imm2)
                }
                patrimonionetto5 = findViewById(R.id.patrimonionetto_5)
                if (patrimonionetto5!!.getText().toString().length != 0) {
                    val patnetto5 = patrimonionetto5!!.getText().toString().toDouble()
                    F10.setCellValue(patnetto5)
                }
                B10.setCellValue("2) Costi  di sviluppo")
                B10.setCellStyle(grassetto)
                E10.setCellValue("V. Riserve statutarie")
                // 11 RIGA
                val row11 = sheet.createRow(10)
                val A11 = row11.createCell(0)
                val B11 = row11.createCell(1)
                val C11 = row11.createCell(2)
                val D11 = row11.createCell(3)
                val E11 = row11.createCell(4)
                val F11 = row11.createCell(5)
                patrimonionetto6 = findViewById(R.id.patrimonionetto_6)
                if (patrimonionetto6!!.getText().toString().length != 0) {
                    val patnetto6 = patrimonionetto6!!.getText().toString().toDouble()
                    F11.setCellValue(patnetto6)
                }
                immobilizzazioni3 = findViewById(R.id.immobilizzazioni_3)
                if (immobilizzazioni3!!.getText().toString().length != 0) {
                    val imm3 = immobilizzazioni3!!.getText().toString().toDouble()
                    C11.setCellValue(imm3)
                }
                B11.setCellValue("3) Diritti di brevetto industriale")
                E11.setCellValue("VI. Altre riserve, distintamente indicate")
                // 12 RIGA
                val row12 = sheet.createRow(11)
                val A12 = row12.createCell(0)
                val B12 = row12.createCell(1)
                val C12 = row12.createCell(2)
                val D12 = row12.createCell(3)
                val E12 = row12.createCell(4)
                val F12 = row12.createCell(5)

                // patrimonionetto7=findViewById(R.id.patrimonionetto_7);
                if (patrimonionetto7.text.toString().length != 0) {
                    val patnetto7 = patrimonionetto7.text.toString().toDouble()
                    F12.setCellValue(patnetto7)
                }
                immobilizzazioni4 = findViewById(R.id.immobilizzazioni_4)
                if (immobilizzazioni4!!.getText().toString().length != 0) {
                    val imm4 = immobilizzazioni4!!.getText().toString().toDouble()
                    C12.setCellValue(imm4)
                }
                B12.setCellValue("4) Concessioni, licenze,  marchi e diritti simili")
                E12.setCellValue("VII. Riserva per operazioni di copertura dei flussi finanziari attesi")
                E12.setCellStyle(grassetto)
                // 13 RIGA
                val row13 = sheet.createRow(12)
                val A13 = row13.createCell(0)
                val B13 = row13.createCell(1)
                val C13 = row13.createCell(2)
                val D13 = row13.createCell(3)
                val E13 = row13.createCell(4)
                val F13 = row13.createCell(5)
                patrimonionetto8 = findViewById(R.id.patrimonionetto_8)
                if (patrimonionetto8!!.getText().toString().length != 0) {
                    val patnetto8 = patrimonionetto8!!.getText().toString().toDouble()
                    F13.setCellValue(patnetto8)
                }
                immobilizzazioni5 = findViewById(R.id.immobilizzazioni_5)
                if (immobilizzazioni5!!.getText().toString().length != 0) {
                    val imm5 = immobilizzazioni5!!.getText().toString().toDouble()
                    C13.setCellValue(imm5)
                }
                B13.setCellValue("5) Avviamento")
                E13.setCellValue("VIII. Utili (perdite) portati a nuovo")
                // 14 RIGA
                val row14 = sheet.createRow(13)
                val A14 = row14.createCell(0)
                val B14 = row14.createCell(1)
                val C14 = row14.createCell(2)
                val D14 = row14.createCell(3)
                val E14 = row14.createCell(4)
                val F14 = row14.createCell(5)
                patrimonionetto9 = findViewById(R.id.patrimonionetto_9)
                if (patrimonionetto9!!.getText().toString().length != 0) {
                    val patnetto9 = patrimonionetto9!!.getText().toString().toDouble()
                    F14.setCellValue(patnetto9)
                }
                immobilizzazioni6 = findViewById(R.id.immobilizzazioni_6)
                fornitoriimmobilizzazioniimmaterialicacconti =
                    findViewById(R.id.fornitori_imm_imm_cacconti)
                if (immobilizzazioni6!!.getText()
                        .toString().length != 0 || fornitoriimmobilizzazioniimmaterialicacconti!!.getText()
                        .toString().length != 0
                ) {
                    val imm6: Double = if (immobilizzazioni6!!.getText().toString().trim { it <= ' ' }
                            .isEmpty()) 0.0 else immobilizzazioni6!!.getText().toString().toDouble()
                    val fornitoriimmcacconti: Double =
                        if (fornitoriimmobilizzazioniimmaterialicacconti!!.getText().toString()
                                .trim { it <= ' ' }
                                .isEmpty()
                        ) 0.0 else fornitoriimmobilizzazioniimmaterialicacconti!!.getText().toString()
                            .toDouble()
                    C14.setCellValue(imm6 + fornitoriimmcacconti)
                }
                B14.setCellValue("6) Immobilizzazioni in corso e acconti")
                E14.setCellValue("IX. Utile (perdita) d'esercizio")

                // 15 RIGA
                val row15 = sheet.createRow(14)
                val A15 = row15.createCell(0)
                val B15 = row15.createCell(1)
                val C15 = row15.createCell(2)
                val D15 = row15.createCell(3)
                val E15 = row15.createCell(4)
                val F15 = row15.createCell(5)

                // patrimonionetto10=findViewById(R.id.patrimonionetto_10);
                if (patrimonionetto10.text.toString().length != 0) {
                    val patnetto10 = patrimonionetto10.text.toString().toDouble()
                    F15.setCellValue(patnetto10)
                }
                immobilizzazioni7 = findViewById(R.id.immobilizzazioni_7)
                if (immobilizzazioni7!!.getText().toString().length != 0) {
                    val imm7 = immobilizzazioni7!!.getText().toString().toDouble()
                    C15.setCellValue(imm7)
                }
                B15.setCellValue("7) Altre")
                E15.setCellValue("X. Riserva negativa per azioni proprie in portafoglio")
                E15.setCellStyle(grassetto)
                // 16 RIGA
                val row16 = sheet.createRow(15)
                val A16 = row16.createCell(0)
                val B16 = row16.createCell(1)
                val C16 = row16.createCell(2)
                val D16 = row16.createCell(3)
                val E16 = row16.createCell(4)
                val F16 = row16.createCell(5)
                A16.setCellValue("II)")
                A16.setCellStyle(grassetto)
                B16.setCellValue("Immobilizzazioni Materiali")
                B16.setCellStyle(grassetto)
                C16.cellFormula = "SUM(C17:C22)"
                C16.setCellStyle(grassetto)
                // 17 RIGA
                val row17 = sheet.createRow(16)
                val A17 = row17.createCell(0)
                val B17 = row17.createCell(1)
                val C17 = row17.createCell(2)
                val D17 = row17.createCell(3)
                val E17 = row17.createCell(4)
                val F17 = row17.createCell(5)
                immobilizzazionimateriali1 = findViewById(R.id.immobilizzazionimateriali_1)
                if (immobilizzazionimateriali1!!.getText().toString().length != 0) {
                    val immM1 = immobilizzazionimateriali1!!.getText().toString().toDouble()
                    C17.setCellValue(immM1)
                }
                B17.setCellValue("1) Terreni e fabbricati")

                //18 RIGA
                val row18 = sheet.createRow(17)
                val A18 = row18.createCell(0)
                val B18 = row18.createCell(1)
                val C18 = row18.createCell(2)
                val D18 = row18.createCell(3)
                val E18 = row18.createCell(4)
                val F18 = row18.createCell(5)
                immobilizzazionimateriali2 = findViewById(R.id.immobilizzazionimateriali_2)
                if (immobilizzazionimateriali2!!.getText().toString().length != 0) {
                    val immM2 = immobilizzazionimateriali2!!.getText().toString().toDouble()
                    C18.setCellValue(immM2)
                }
                B18.setCellValue("2) Impianti e macchinari")
                D18.setCellValue("B)")
                D18.setCellStyle(grassetto)
                E18.setCellValue("FONDI PER RISCHI E ONERI")
                E18.setCellStyle(grassetto)
                F18.cellFormula = "SUM(F19:F24)"
                F18.setCellStyle(grassetto)
                // 19 RIGA
                val row19 = sheet.createRow(18)
                val A19 = row19.createCell(0)
                val B19 = row19.createCell(1)
                val C19 = row19.createCell(2)
                val D19 = row19.createCell(3)
                val E19 = row19.createCell(4)
                val F19 = row19.createCell(5)
                immobilizzazionimateriali3 = findViewById(R.id.immobilizzazionimateriali_3)
                if (immobilizzazionimateriali3!!.getText().toString().length != 0) {
                    val immM3 = immobilizzazionimateriali3!!.getText().toString().toDouble()
                    C19.setCellValue(immM3)
                }
                fondiperrischieoneri1 = findViewById(R.id.fondiperrischieoneri_1)
                if (fondiperrischieoneri1!!.getText().toString().length != 0) {
                    val fpr1 = fondiperrischieoneri1!!.getText().toString().toDouble()
                    F19.setCellValue(fpr1)
                }
                B19.setCellValue("3) Attrezzature industriali e commerciali")
                E19.setCellValue("1) Per trattamento di quiescenza e obblighi simili")
                // 20 RIGA
                val row20 = sheet.createRow(19)
                val A20 = row20.createCell(0)
                val B20 = row20.createCell(1)
                val C20 = row20.createCell(2)
                val D20 = row20.createCell(3)
                val E20 = row20.createCell(4)
                val F20 = row20.createCell(5)
                immobilizzazionimaterialifa = findViewById(R.id.immobilizzazionimateriali_F_A)
                if (immobilizzazionimaterialifa!!.getText().toString().length != 0) {
                    val immMfa = immobilizzazionimaterialifa!!.getText().toString().toDouble()
                    C20.setCellValue(immMfa)
                }
                B20.setCellValue("F.Amm. Edifici industriali")
                E20.setCellValue("2) Per imposte")
                fondiperrischieoneri2 = findViewById(R.id.fondiperrischieoneri_2)
                if (fondiperrischieoneri2!!.getText().toString().length != 0) {
                    val fpr2 = fondiperrischieoneri2!!.getText().toString().toDouble()
                    F20.setCellValue(fpr2)
                }

                //21 RIGA
                val row21 = sheet.createRow(20)
                val A21 = row21.createCell(0)
                val B21 = row21.createCell(1)
                val C21 = row21.createCell(2)
                val D21 = row21.createCell(3)
                val E21 = row21.createCell(4)
                val F21 = row21.createCell(5)
                immobilizzazionimateriali4 = findViewById(R.id.immobilizzazionimateriali_4)
                imballaggidurevoli = findViewById(R.id.imballaggidurevoli)
                if (immobilizzazionimateriali4!!.getText()
                        .toString().length != 0 || imballaggidurevoli!!.getText()
                        .toString().length != 0
                ) {
                    val immM4: Double =
                        if (immobilizzazionimateriali4!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else immobilizzazionimateriali4!!.getText().toString()
                            .toDouble()
                    val imballaggi: Double =
                        if (imballaggidurevoli!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else imballaggidurevoli!!.getText().toString()
                            .toDouble()
                    C21.setCellValue(immM4 + imballaggi)
                }
                B21.setCellValue("4)Altri beni")
                E21.setCellValue("3) strumenti finanziari derivati passivi")
                E21.setCellStyle(grassetto)
                fondiperrischieoneri3 = findViewById(R.id.fondiperrischieoneri_3)
                if (fondiperrischieoneri3!!.getText().toString().length != 0) {
                    val fpr3 = fondiperrischieoneri3!!.getText().toString().toDouble()
                    F21.setCellValue(fpr3)
                }

                // 22 RIGA
                val row22 = sheet.createRow(21)
                val A22 = row22.createCell(0)
                val B22 = row22.createCell(1)
                val C22 = row22.createCell(2)
                val D22 = row22.createCell(3)
                val E22 = row22.createCell(4)
                val F22 = row22.createCell(5)
                immobilizzazionimateriali5 = findViewById(R.id.immobilizzazionimateriali_5)
                fornitoriimmobilizzazioniimmaterialicacconti =
                    findViewById(R.id.fornitori_imm_mat_cacconti)
                if (immobilizzazionimateriali5!!.getText()
                        .toString().length != 0 || fornitoriimmobilizzazioniimmaterialicacconti!!.getText()
                        .toString().length != 0
                ) {
                    val immM5: Double =
                        if (immobilizzazionimateriali5!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else immobilizzazionimateriali5!!.getText().toString()
                            .toDouble()
                    val immobilizzazionimat: Double =
                        if (fornitoriimmobilizzazioniimmaterialicacconti!!.getText().toString()
                                .trim { it <= ' ' }
                                .isEmpty()
                        ) 0.0 else fornitoriimmobilizzazioniimmaterialicacconti!!.getText().toString()
                            .toDouble()
                    C22.setCellValue(immM5 + immobilizzazionimat)
                }
                fondiperrischieoneri4 = findViewById(R.id.fondiperrischieoneri_4)
                if (fondiperrischieoneri4!!.getText().toString().length != 0) {
                    val fpr4 = fondiperrischieoneri4!!.getText().toString().toDouble()
                    F22.setCellValue(fpr4)
                }
                B22.setCellValue("5) Immobilizzazioni in corso e acconti")
                E22.setCellValue("4) Altri")
                // 23 RIGA
                val row23 = sheet.createRow(22)
                val A23 = row23.createCell(0)
                val B23 = row23.createCell(1)
                val C23 = row23.createCell(2)
                val D23 = row23.createCell(3)
                val E23 = row23.createCell(4)
                val F23 = row23.createCell(5)
                A23.setCellValue("III)")
                A23.setCellStyle(grassetto)
                B23.setCellValue("Immobilizzazioni Finanziarie")
                B23.setCellStyle(grassetto)
                C23.cellFormula = "SUM(C24:C27)"
                C23.setCellStyle(grassetto)

                // 24 RIGA
                val row24 = sheet.createRow(23)
                val A24 = row24.createCell(0)
                val B24 = row24.createCell(1)
                val C24 = row24.createCell(2)
                val D24 = row24.createCell(3)
                val E24 = row24.createCell(4)
                val F24 = row24.createCell(5)
                B24.setCellValue("1) Partecipazioni")
                immobilizzazionifinanziarie1 = findViewById(R.id.immobilizzazionifinanziarie_1)
                if (immobilizzazionifinanziarie1!!.getText().toString().length != 0) {
                    val immf1 = immobilizzazionifinanziarie1!!.getText().toString().toDouble()
                    C24.setCellValue(immf1)
                }


                // 25 RIGA
                val row25 = sheet.createRow(24)
                val A25 = row25.createCell(0)
                val B25 = row25.createCell(1)
                val C25 = row25.createCell(2)
                val D25 = row25.createCell(3)
                val E25 = row25.createCell(4)
                val F25 = row25.createCell(5)
                B25.setCellValue("2) Crediti")
                D25.setCellValue("C)")
                D25.setCellStyle(grassetto)
                E25.setCellValue("TFR")
                E25.setCellStyle(grassetto)
                F25.setCellStyle(grassetto)
                mutuiattivi = findViewById(R.id.mutui_attivi)
                immobilizzazionifinanziarie2 = findViewById(R.id.immobilizzazionifinanziarie_2)
                if (immobilizzazionifinanziarie2!!.getText()
                        .toString().length != 0 || mutuiattivi!!.getText().toString().length != 0
                ) {
                    val immf2: Double =
                        if (immobilizzazionifinanziarie2!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else immobilizzazionifinanziarie2!!.getText().toString()
                            .toDouble()
                    val mutuiattivii: Double =
                        if (mutuiattivi!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else mutuiattivi!!.getText().toString().toDouble()
                    C25.setCellValue(immf2 + mutuiattivii)
                }
                tfr1 = findViewById(R.id.tfr_1)
                if (tfr1!!.getText().toString().length != 0) {
                    val tfr = tfr1!!.getText().toString().toDouble()
                    F25.setCellValue(tfr)
                }

                // 26 RIGA
                val row26 = sheet.createRow(25)
                val A26 = row26.createCell(0)
                val B26 = row26.createCell(1)
                val C26 = row26.createCell(2)
                val D26 = row26.createCell(3)
                val E26 = row26.createCell(4)
                val F26 = row26.createCell(5)
                B26.setCellValue("3) Altri Titoli ")
                immobilizzazionifinanziarie3 = findViewById(R.id.immobilizzazionifinanziarie_3)
                if (immobilizzazionifinanziarie3!!.getText().toString().length != 0) {
                    val immf3 = immobilizzazionifinanziarie3!!.getText().toString().toDouble()
                    C26.setCellValue(immf3)
                }

                // 27 RIGA
                val row27 = sheet.createRow(26)
                val A27 = row27.createCell(0)
                val B27 = row27.createCell(1)
                val C27 = row27.createCell(2)
                val D27 = row27.createCell(3)
                val E27 = row27.createCell(4)
                val F27 = row27.createCell(5)
                B27.setCellValue("4) Strumenti finanziari derivati attivi")
                B27.setCellStyle(grassetto)
                immobilizzazionifinanziarie4 = findViewById(R.id.immobilizzazionifinanziarie_4)
                if (immobilizzazionifinanziarie4!!.getText().toString().length != 0) {
                    val immf4 = immobilizzazionifinanziarie4!!.getText().toString().toDouble()
                    C27.setCellValue(immf4)
                }

                // 28 RIGA
                val row28 = sheet.createRow(27)
                val A28 = row28.createCell(0)
                val B28 = row28.createCell(1)
                val C28 = row28.createCell(2)
                val D28 = row28.createCell(3)
                val E28 = row28.createCell(4)
                val F28 = row28.createCell(5)
                A28.setCellValue("C)")
                A28.setCellStyle(grassetto)
                B28.setCellValue("ATTIVO CIRCOLANTE")
                B28.setCellStyle(grassetto)
                C28.cellFormula = "C29+C35+C55+C46"
                C28.setCellStyle(grassetto)
                D28.setCellValue("D)")
                D28.setCellStyle(grassetto)
                E28.setCellValue("DEBITI")
                E28.setCellStyle(grassetto)
                F28.cellFormula = "F29+F30+F31+F32+F35+F38+F39+F40+F41+F42+F43+F44+F45+F46+F47"
                F28.setCellStyle(grassetto)
                // 29 RIGA
                val row29 = sheet.createRow(28)
                val A29 = row29.createCell(0)
                val B29 = row29.createCell(1)
                val C29 = row29.createCell(2)
                val D29 = row29.createCell(3)
                val E29 = row29.createCell(4)
                val F29 = row29.createCell(5)
                A29.setCellValue("I)")
                A29.setCellStyle(grassetto)
                B29.setCellValue("Rimanenze")
                B29.setCellStyle(grassetto)
                C29.cellFormula = "SUM(C30:C34)"
                C29.setCellStyle(grassetto)
                E29.setCellValue("1) Obbligazioni ")
                debiti1 = findViewById(R.id.debiti_1)
                if (debiti1!!.getText().toString().length != 0) {
                    val debt1 = debiti1!!.getText().toString().toDouble()
                    F29.setCellValue(debt1)
                }

                // 30 RIGA
                val row30 = sheet.createRow(29)
                val A30 = row30.createCell(0)
                val B30 = row30.createCell(1)
                val C30 = row30.createCell(2)
                val D30 = row30.createCell(3)
                val E30 = row30.createCell(4)
                val F30 = row30.createCell(5)
                B30.setCellValue("1) Materie prime, sussidiarie e di consumo")
                E30.setCellValue("2) Obbligazioni convertibili")
                rimanenze1 = findViewById(R.id.rimanenze_1)
                if (rimanenze1!!.getText().toString().length != 0) {
                    val rim1 = rimanenze1!!.getText().toString().toDouble()
                    C30.setCellValue(rim1)
                }
                debiti2 = findViewById(R.id.debiti_2)
                if (debiti2!!.getText().toString().length != 0) {
                    val debt2 = debiti2!!.getText().toString().toDouble()
                    F30.setCellValue(debt2)
                }

                // 31 RIGA
                val row31 = sheet.createRow(30)
                val A31 = row31.createCell(0)
                val B31 = row31.createCell(1)
                val C31 = row31.createCell(2)
                val D31 = row31.createCell(3)
                val E31 = row31.createCell(4)
                val F31 = row31.createCell(5)
                B31.setCellValue("2) Prodotti in corso di lavorazione e semilavorati")
                E31.setCellValue("3) Debiti verso soci per finanziamenti")
                rimanenze2 = findViewById(R.id.rimanenze_2)
                if (rimanenze2!!.getText().toString().length != 0) {
                    val rim2 = rimanenze2!!.getText().toString().toDouble()
                    C31.setCellValue(rim2)
                }
                debiti3 = findViewById(R.id.debiti_3)
                if (debiti3!!.getText().toString().length != 0) {
                    val debt3 = debiti3!!.getText().toString().toDouble()
                    F31.setCellValue(debt3)
                }

                // 32 RIGA
                val row32 = sheet.createRow(31)
                val A32 = row32.createCell(0)
                val B32 = row32.createCell(1)
                val C32 = row32.createCell(2)
                val D32 = row32.createCell(3)
                val E32 = row32.createCell(4)
                val F32 = row32.createCell(5)
                B32.setCellValue("3) Lavori in corso su ordinazione")
                E32.setCellValue("4) Debiti verso banche:")
                F32.cellFormula = "F33+F34"
                rimanenze3 = findViewById(R.id.rimanenze_3)
                if (rimanenze3!!.getText().toString().length != 0) {
                    val rim3 = rimanenze3!!.getText().toString().toDouble()
                    C32.setCellValue(rim3)
                }
                // 33 RIGA
                val row33 = sheet.createRow(32)
                val A33 = row33.createCell(0)
                val B33 = row33.createCell(1)
                val C33 = row33.createCell(2)
                val D33 = row33.createCell(3)
                val E33 = row33.createCell(4)
                val F33 = row33.createCell(5)
                B33.setCellValue("4) Prodotti finiti")
                E33.setCellValue("      - entro l'anno")
                rimanenze4 = findViewById(R.id.rimanenze_4)
                if (rimanenze4!!.getText().toString().length != 0) {
                    val rim4 = rimanenze4!!.getText().toString().toDouble()
                    C33.setCellValue(rim4)
                }
                debiti4punto1 = findViewById(R.id.debiti_4_1)
                mutuipassivi = findViewById(R.id.mutui_passivi)
                val mutuipassiviText = mutuipassivi!!.getText().toString().trim { it <= ' ' }
                val debiti4punto1text = debiti4punto1!!.getText().toString().trim { it <= ' ' }
                if (!mutuipassiviText.isEmpty() || !debiti4punto1text.isEmpty()) {
                    try {
                        val mutuipassivii: Double =
                            if (mutuipassiviText.isEmpty()) 0.0 else mutuipassiviText.toDouble()
                        val debt4punto1: Double =
                            if (debiti4punto1text.isEmpty()) 0.0 else debiti4punto1text.toDouble()
                        F33.setCellValue(debt4punto1 + mutuipassivii)
                    } catch (e: Exception) {
                    }
                }
                // 34 RIGA
                val row34 = sheet.createRow(33)
                val A34 = row34.createCell(0)
                val B34 = row34.createCell(1)
                val C34 = row34.createCell(2)
                val D34 = row34.createCell(3)
                val E34 = row34.createCell(4)
                val F34 = row34.createCell(5)
                B34.setCellValue("5) Acconti")
                E34.setCellValue("      - oltre l'anno")
                rimanenze5 = findViewById(R.id.rimanenze_5)
                fornitoriacconti = findViewById(R.id.fornitori_cacconti)
                if (rimanenze5!!.getText()
                        .toString().length != 0 || fornitoriacconti.toString().length != 0
                ) {
                    val rim5: Double = if (rimanenze5!!.getText().toString().trim { it <= ' ' }
                            .isEmpty()) 0.0 else rimanenze5!!.getText().toString().toDouble()
                    val fornitoricacc: Double =
                        if (fornitoriacconti!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else fornitoriacconti!!.getText().toString().toDouble()
                    C34.setCellValue(rim5 + fornitoricacc)
                }
                debiti4punto2 = findViewById(R.id.debiti_4_2)
                if (debiti4punto2!!.getText().toString().length != 0) {
                    val debt4punto2 = debiti4punto2!!.getText().toString().toDouble()
                    F34.setCellValue(debt4punto2)
                }

                // 35 RIGA
                val row35 = sheet.createRow(34)
                val A35 = row35.createCell(0)
                val B35 = row35.createCell(1)
                val C35 = row35.createCell(2)
                val D35 = row35.createCell(3)
                val E35 = row35.createCell(4)
                val F35 = row35.createCell(5)
                A35.setCellValue("II)")
                A35.setCellStyle(grassetto)
                B35.setCellValue("Crediti")
                B35.setCellStyle(grassetto)
                C35.cellFormula = "C36+C39+C40+C41+C42+C43+C44+C45"
                C35.setCellStyle(grassetto)
                E35.setCellValue("5) Debiti verso altri finanziatori:")
                F35.cellFormula = "F36+F37"
                // 36 RIGA
                val row36 = sheet.createRow(35)
                val A36 = row36.createCell(0)
                val B36 = row36.createCell(1)
                val C36 = row36.createCell(2)
                val D36 = row36.createCell(3)
                val E36 = row36.createCell(4)
                val F36 = row36.createCell(5)
                B36.setCellValue("1) Verso clienti")
                fatturedaemettere = findViewById(R.id.fatturedaemettere)
                crediticommdiversi = findViewById(R.id.crediticommdiversi)
                debiti5punto1 = findViewById(R.id.debiti_5_1)
                crediti1punto1 = findViewById(R.id.crediti_1_1)
                crediti1punto2 = findViewById(R.id.crediti_1_2)
                clienticostianticipati = findViewById(R.id.clienticcostianticipati)
                cambialiattive = findViewById(R.id.crediti_cambiali_attive)
                cambialiallosconto = findViewById(R.id.cambialiallosconto)
                cambialiallincasso = findViewById(R.id.cambialiallincasso)
                cambialiinsolute = findViewById(R.id.cambialiinsolute)
                fondorischisucrediti = findViewById(R.id.fondorischisucrediti)
                fondosvalutazionecrediti = findViewById(R.id.fondosvalutazionicrediti)
                if ((crediti1punto2!!.getText().toString().length != 0) || (crediti1punto1!!.getText()
                        .toString().length != 0) || (crediticommdiversi!!.getText()
                        .toString().length != 0) || (fatturedaemettere!!.getText()
                        .toString().length != 0) || (clienticostianticipati!!.getText()
                        .toString().length != 0) || (cambialiattive!!.getText()
                        .toString().length != 0) || (cambialiallosconto!!.getText()
                        .toString().length != 0) || (cambialiallincasso!!.getText()
                        .toString().length != 0) || (cambialiinsolute!!.getText()
                        .toString().length != 0) || (fondorischisucrediti!!.getText()
                        .toString().length != 0) || (fondosvalutazionecrediti!!.getText()
                        .toString().length != 0)
                ) {
                    val cr1punto2: Double =
                        if (crediti1punto2!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else crediti1punto2!!.getText().toString().toDouble()
                    val cr1punto1: Double =
                        if (crediti1punto1!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else crediti1punto1!!.getText().toString().toDouble()
                    val fatturedaemetter: Double =
                        if (fatturedaemettere!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else fatturedaemettere!!.getText().toString().toDouble()
                    val crediticomdiv: Double =
                        if (crediticommdiversi!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else crediticommdiversi!!.getText().toString()
                            .toDouble()
                    val clienticosti: Double =
                        if (clienticostianticipati!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else clienticostianticipati!!.getText().toString()
                            .toDouble()
                    val cambialiattiv: Double =
                        if (cambialiattive!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else cambialiattive!!.getText().toString().toDouble()
                    val cambialialloscont: Double =
                        if (cambialiallosconto!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else cambialiallosconto!!.getText().toString()
                            .toDouble()
                    val cambialiallincass: Double =
                        if (cambialiallincasso!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else cambialiallincasso!!.getText().toString()
                            .toDouble()
                    val cambialiinsolut: Double =
                        if (cambialiinsolute!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else cambialiinsolute!!.getText().toString().toDouble()
                    val fondorischisucredit: Double =
                        if (fondorischisucrediti!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else fondorischisucrediti!!.getText().toString()
                            .toDouble()
                    val fondosvalutazionecredit: Double =
                        if (fondosvalutazionecrediti!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else fondosvalutazionecrediti!!.getText().toString()
                            .toDouble()
                    C36.setCellValue(cr1punto2 + cr1punto1 + fatturedaemetter + crediticomdiv + clienticosti + cambialiattiv + cambialialloscont + cambialiallincass + cambialiinsolut + fondorischisucredit + fondosvalutazionecredit)
                }
                E36.setCellValue("- entro l'anno")
                if (debiti5punto1!!.getText().toString().length != 0) {
                    val debt5punto1: Double =
                        if (debiti5punto1!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else debiti5punto1!!.getText().toString().toDouble()
                    F36.setCellValue(debt5punto1)
                }


                // 37 RIGA
                val row37 = sheet.createRow(36)
                val A37 = row37.createCell(0)
                val B37 = row37.createCell(1)
                val C37 = row37.createCell(2)
                val D37 = row37.createCell(3)
                val E37 = row37.createCell(4)
                val F37 = row37.createCell(5)
                B37.setCellValue("- entro l'anno")
                E37.setCellValue("- oltre l'anno")
                if (crediti1punto1!!.getText().toString().length != 0) {
                    val cr1punto1: Double =
                        if (crediti1punto1!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else crediti1punto1!!.getText().toString().toDouble()
                    C37.setCellValue(cr1punto1)
                }
                debiti5punto2 = findViewById(R.id.debiti_5_2)
                if (debiti5punto2!!.getText().toString().length != 0) {
                    val debt5punto2 = debiti5punto2!!.getText().toString().toDouble()
                    F37.setCellValue(debt5punto2)
                }

                //38 RIGA
                val row38 = sheet.createRow(37)
                val A38 = row38.createCell(0)
                val B38 = row38.createCell(1)
                val C38 = row38.createCell(2)
                val D38 = row38.createCell(3)
                val E38 = row38.createCell(4)
                val F38 = row38.createCell(5)
                B38.setCellValue("- oltre l'anno")
                E38.setCellValue("6) Acconti")
                clienticacconti = findViewById(R.id.clienti_cacconti)
                if (crediti1punto2!!.getText().toString().length != 0) {
                    val cr1punto2: Double =
                        if (crediti1punto2!!.getText().toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else crediti1punto2!!.getText().toString().toDouble()
                    C38.setCellValue(cr1punto2)
                }
                debiti6 = findViewById(R.id.debiti_6)
                if (debiti6!!.getText().toString().length != 0 || clienticacconti!!.getText()
                        .toString().length != 0
                ) {
                    val debt6: Double = if (debiti6!!.getText().toString().trim { it <= ' ' }
                            .isEmpty()) 0.0 else debiti6!!.getText().toString().toDouble()
                    val clienticacconti: Double =
                        if (clientiacconti!!.text.toString().trim { it <= ' ' }
                                .isEmpty()) 0.0 else clientiacconti!!.text.toString().toDouble()
                    F38.setCellValue(debt6 + clienticacconti)
                }

                // 39 RIGA
                val row39 = sheet.createRow(38)
                val A39 = row39.createCell(0)
                val B39 = row39.createCell(1)
                val C39 = row39.createCell(2)
                val D39 = row39.createCell(3)
                val E39 = row39.createCell(4)
                val F39 = row39.createCell(5)
                crediti2 = findViewById(R.id.crediti_2)
                val crediti2Text = crediti2!!.getText().toString().trim { it <= ' ' }
                if (!crediti2Text.isEmpty()) {
                    try {
                        val cr2: Double = if (crediti2Text.isEmpty()) 0.0 else crediti2Text.toDouble()
                        C39.setCellValue(cr2)
                    } catch (e: Exception) {
                    }
                }
                fatturedaricevere = findViewById(R.id.debiti_fatture_da_ricevere)
                debiti7 = findViewById(R.id.debiti_7)
                val fatturedaricevereText =
                    fatturedaricevere!!.getText().toString().trim { it <= ' ' }
                val debiti7Text = debiti7!!.getText().toString().trim { it <= ' ' }
                if (!fatturedaricevereText.isEmpty() || !debiti7Text.isEmpty()) {
                    try {
                        val debt7: Double =
                            if (debiti7Text.isEmpty()) 0.0 else debiti7!!.getText().toString()
                                .toDouble()
                        val fatturedariceveree: Double =
                            if (fatturedaricevereText.isEmpty()) 0.0 else fatturedaricevereText.toDouble()
                        F39.setCellValue(debt7 + fatturedariceveree)
                    } catch (e: Exception) {
                    }
                }
                B39.setCellValue("2) Verso imprese controllate")
                E39.setCellValue("7) Debiti verso fornitori")
                // 40 RIGA
                val row40 = sheet.createRow(39)
                val A40 = row40.createCell(0)
                val B40 = row40.createCell(1)
                val C40 = row40.createCell(2)
                val D40 = row40.createCell(3)
                val E40 = row40.createCell(4)
                val F40 = row40.createCell(5)
                crediti3 = findViewById(R.id.crediti_3)
                if (crediti3!!.getText().toString().length != 0) {
                    val cr3 = crediti3!!.getText().toString().toDouble()
                    C40.setCellValue(cr3)
                }
                debiti8 = findViewById(R.id.debiti_8)
                if (debiti8!!.getText().toString().length != 0) {
                    val debt8 = debiti8!!.getText().toString().toDouble()
                    F40.setCellValue(debt8)
                }
                B40.setCellValue("3) Verso imprese collegate")
                E40.setCellValue("8) Debiti rappresentati da titoli di credito")
                val row41 = sheet.createRow(40)
                val A41 = row41.createCell(0)
                val B41 = row41.createCell(1)
                val C41 = row41.createCell(2)
                val D41 = row41.createCell(3)
                val E41 = row41.createCell(4)
                val F41 = row41.createCell(5)
                crediti4 = findViewById(R.id.crediti_4)
                if (crediti4!!.getText().toString().length != 0) {
                    val cr4 = crediti4!!.getText().toString().toDouble()
                    C41.setCellValue(cr4)
                }
                debiti9 = findViewById(R.id.debiti_9)
                if (debiti9!!.getText().toString().length != 0) {
                    val debt9 = debiti9!!.getText().toString().toDouble()
                    F41.setCellValue(debt9)
                }
                B41.setCellValue("4) Verso controllanti")
                E41.setCellValue("9) Debiti verso imprese controllate")
                // 42 RIGA
                val row42 = sheet.createRow(41)
                val A42 = row42.createCell(0)
                val B42 = row42.createCell(1)
                val C42 = row42.createCell(2)
                val D42 = row42.createCell(3)
                val E42 = row42.createCell(4)
                val F42 = row42.createCell(5)

                // crediti5=findViewById(R.id.crediti_5);
                if (crediti5.text.toString().length != 0) {
                    val cr5 = crediti5.text.toString().toDouble()
                    C42.setCellValue(cr5)
                }
                debiti10 = findViewById(R.id.debiti_10)
                if (debiti10!!.getText().toString().length != 0) {
                    val debt10 = debiti10!!.getText().toString().toDouble()
                    F42.setCellValue(debt10)
                }
                B42.setCellValue("5) Verso imprese sottoposte al controllo delle controllanti")
                B42.setCellStyle(grassetto)
                E42.setCellValue("10) Debiti verso imprese collegate")
                // 43 RIGA
                val row43 = sheet.createRow(42)
                val A43 = row43.createCell(0)
                val B43 = row43.createCell(1)
                val C43 = row43.createCell(2)
                val D43 = row43.createCell(3)
                val E43 = row43.createCell(4)
                val F43 = row43.createCell(5)
                crediti5punto2 = findViewById(R.id.crediti_5_2)
                if (crediti5punto2!!.getText().toString().length != 0) {
                    val cr5punto2 = crediti5punto2!!.getText().toString().toDouble()
                    C43.setCellValue(cr5punto2)
                }
                debiti11 = findViewById(R.id.debiti_11)
                if (debiti11!!.getText().toString().length != 0) {
                    val debt11 = debiti11!!.getText().toString().toDouble()
                    F43.setCellValue(debt11)
                }
                B43.setCellValue("5bis)crediti tributari")
                E43.setCellValue("11) Debiti verso controllanti")
                //44 RIGA
                val row44 = sheet.createRow(43)
                val A44 = row44.createCell(0)
                val B44 = row44.createCell(1)
                val C44 = row44.createCell(2)
                val D44 = row44.createCell(3)
                val E44 = row44.createCell(4)
                val F44 = row44.createCell(5)
                crediti5punto3 = findViewById(R.id.crediti_5_3)
                if (crediti5punto3!!.getText().toString().length != 0) {
                    val cr5punto3 = crediti5punto3!!.getText().toString().toDouble()
                    C44.setCellValue(cr5punto3)
                }

                // debiti11punto2=findViewById(R.id.debiti_11_2);
                if (debiti11punto2.text.toString().length != 0) {
                    val debt11punto2 = debiti11punto2.text.toString().toDouble()
                    F44.setCellValue(debt11punto2)
                }
                B44.setCellValue("5ter)imposte anticipate")
                E44.setCellValue("11bis) debiti verso imprese sottoposte al controllo delle controllanti")
                //45 RIGA
                val row45 = sheet.createRow(44)
                val A45 = row45.createCell(0)
                val B45 = row45.createCell(1)
                val C45 = row45.createCell(2)
                val D45 = row45.createCell(3)
                val E45 = row45.createCell(4)
                val F45 = row45.createCell(5)
                crediti5punto4 = findViewById(R.id.crediti_5_4)
                if (crediti5punto4!!.getText().toString().length != 0) {
                    val cr5punto4 = crediti5punto4!!.getText().toString().toDouble()
                    C45.setCellValue(cr5punto4)
                }
                debiti12 = findViewById(R.id.debiti_12)
                erariociv = findViewById(R.id.erario_civa)
                if (debiti12!!.getText().toString().length != 0 || erariociv!!.getText()
                        .toString().length != 0
                ) {
                    val debt12: Double = if (debiti12!!.getText().toString().trim { it <= ' ' }
                            .isEmpty()) 0.0 else debiti12!!.getText().toString().toDouble()
                    val erariociva: Double = if (erariociv!!.getText().toString().trim { it <= ' ' }
                            .isEmpty()) 0.0 else erariociv!!.getText().toString().toDouble()
                    F45.setCellValue(debt12 + erariociva)
                }
                B45.setCellValue("5quater)verso altri")
                E45.setCellValue("12) Debiti tributari")
                //46 RIGA
                val row46 = sheet.createRow(45)
                val A46 = row46.createCell(0)
                val B46 = row46.createCell(1)
                val C46 = row46.createCell(2)
                val D46 = row46.createCell(3)
                val E46 = row46.createCell(4)
                val F46 = row46.createCell(5)
                B46.setCellValue("III. Attività finan. non immob.")
                B46.setCellStyle(grassetto)
                E46.setCellValue("13) Debiti verso istituti di previdenza e di sicurezza sociale")
                C46.cellFormula = "SUM(C47:C54)"
                C46.setCellStyle(grassetto)


                //   debiti13=findViewById(R.id.debiti_13);
                if (debiti13.text.toString().length != 0) {
                    val debt13 = debiti13.text.toString().toDouble()
                    F46.setCellValue(debt13)
                }
                //47 RIGA
                val row47 = sheet.createRow(46)
                val A47 = row47.createCell(0)
                val B47 = row47.createCell(1)
                val C47 = row47.createCell(2)
                val D47 = row47.createCell(3)
                val E47 = row47.createCell(4)
                val F47 = row47.createCell(5)
                debiti14 = findViewById(R.id.debiti_14)
                if (debiti14!!.getText().toString().length != 0) {
                    val debt14 = debiti14!!.getText().toString().toDouble()
                    F47.setCellValue(debt14)
                }
                immobilizzazionifinanziarienonimmobilizzate1 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_1)
                if (immobilizzazionifinanziarienonimmobilizzate1!!.getText().toString().length != 0) {
                    val immfni1 =
                        immobilizzazionifinanziarienonimmobilizzate1!!.getText().toString().toDouble()
                    C47.setCellValue(immfni1)
                }
                B47.setCellValue("1) Partecipazioni in imprese controllate")
                E47.setCellValue("14) Altri debiti")
                //48 RIGA
                val row48 = sheet.createRow(47)
                val A48 = row48.createCell(0)
                val B48 = row48.createCell(1)
                val C48 = row48.createCell(2)
                val D48 = row48.createCell(3)
                val E48 = row48.createCell(4)
                val F48 = row48.createCell(5)
                B48.setCellValue("2) Partecipazioni in imprese collegate")
                immobilizzazionifinanziarienonimmobilizzate2 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_2)
                if (immobilizzazionifinanziarienonimmobilizzate2!!.getText().toString().length != 0) {
                    val immfni2 =
                        immobilizzazionifinanziarienonimmobilizzate2!!.getText().toString().toDouble()
                    C48.setCellValue(immfni2)
                }
                //49 RIGA
                val row49 = sheet.createRow(48)
                val A49 = row49.createCell(0)
                val B49 = row49.createCell(1)
                val C49 = row49.createCell(2)
                val D49 = row49.createCell(3)
                val E49 = row49.createCell(4)
                val F49 = row49.createCell(5)
                B49.setCellValue("3) Partecipazioni in imprese controllanti")
                immobilizzazionifinanziarienonimmobilizzate3 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_3)
                if (immobilizzazionifinanziarienonimmobilizzate3!!.getText().toString().length != 0) {
                    val immfni3 =
                        immobilizzazionifinanziarienonimmobilizzate3!!.getText().toString().toDouble()
                    C49.setCellValue(immfni3)
                }
                // 50 RIGA
                val row50 = sheet.createRow(49)
                val A50 = row50.createCell(0)
                val B50 = row50.createCell(1)
                val C50 = row50.createCell(2)
                val D50 = row50.createCell(3)
                val E50 = row50.createCell(4)
                val F50 = row50.createCell(5)
                B50.setCellValue("3bis) partecipazioni in imprese sottoposte al controllo delle controllanti")
                B50.setCellStyle(grassetto)


                //  immobilizzazionifinanziarienonimmobilizzate3punto2=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_3_2);
                if (immobilizzazionifinanziarienonimmobilizzate3punto2.text.toString().length != 0) {
                    val immfni3punto2 =
                        immobilizzazionifinanziarienonimmobilizzate3punto2.text.toString()
                            .toDouble()
                    C50.setCellValue(immfni3punto2)
                }
                //51 RIGA
                val row51 = sheet.createRow(50)
                val A51 = row51.createCell(0)
                val B51 = row51.createCell(1)
                val C51 = row51.createCell(2)
                val D51 = row51.createCell(3)
                val E51 = row51.createCell(4)
                val F51 = row51.createCell(5)
                B51.setCellValue("4) altre partecipazioni")
                immobilizzazionifinanziarienonimmobilizzate4 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_4)
                if (immobilizzazionifinanziarienonimmobilizzate4!!.getText().toString().length != 0) {
                    val immfni4 =
                        immobilizzazionifinanziarienonimmobilizzate4!!.getText().toString().toDouble()
                    C51.setCellValue(immfni4)
                }
                //52 RIGA
                val row52 = sheet.createRow(51)
                val A52 = row52.createCell(0)
                val B52 = row52.createCell(1)
                val C52 = row52.createCell(2)
                val D52 = row52.createCell(3)
                val E52 = row52.createCell(4)
                val F52 = row52.createCell(5)
                B52.setCellValue("5) Strumenti finanziari derivati attivi")
                B52.setCellStyle(grassetto)
                immobilizzazionifinanziarienonimmobilizzate5 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_5)
                if (immobilizzazionifinanziarienonimmobilizzate5!!.getText().toString().length != 0) {
                    val immfni5 =
                        immobilizzazionifinanziarienonimmobilizzate5!!.getText().toString().toDouble()
                    C52.setCellValue(immfni5)
                }
                //53 RIGA
                val row53 = sheet.createRow(52)
                val A53 = row53.createCell(0)
                val B53 = row53.createCell(1)
                val C53 = row53.createCell(2)
                val D53 = row53.createCell(3)
                val E53 = row53.createCell(4)
                val F53 = row53.createCell(5)
                B53.setCellValue("6) Altri titoli")
                val row54 = sheet.createRow(53)
                immobilizzazionifinanziarienonimmobilizzate6 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_6)
                if (immobilizzazionifinanziarienonimmobilizzate6!!.getText().toString().length != 0) {
                    val immfni6 =
                        immobilizzazionifinanziarienonimmobilizzate6!!.getText().toString().toDouble()
                    C53.setCellValue(immfni6)
                }

                //54 RIGA
                val A54 = row54.createCell(0)
                val B54 = row54.createCell(1)
                val C54 = row54.createCell(2)
                val D54 = row54.createCell(3)
                val E54 = row54.createCell(4)
                val F54 = row54.createCell(5)
                B54.setCellValue("7) Altre")
                immobilizzazionifinanziarienonimmobilizzate7 =
                    findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_7)
                if (immobilizzazionifinanziarienonimmobilizzate7!!.getText().toString().length != 0) {
                    val immfni7 =
                        immobilizzazionifinanziarienonimmobilizzate7!!.getText().toString().toDouble()
                    C54.setCellValue(immfni7)
                }
                //55 RIGA
                val row55 = sheet.createRow(54)
                val A55 = row55.createCell(0)
                val B55 = row55.createCell(1)
                val C55 = row55.createCell(2)
                val D55 = row55.createCell(3)
                val E55 = row55.createCell(4)
                val F55 = row55.createCell(5)
                A55.setCellValue("IV)")
                A55.setCellStyle(grassetto)
                B55.setCellValue("Disponibilità liquide")
                B55.setCellStyle(grassetto)
                C55.cellFormula = "SUM(C56:C58)"
                C55.setCellStyle(grassetto)
                //56 RIGA
                val row56 = sheet.createRow(55)
                val A56 = row56.createCell(0)
                val B56 = row56.createCell(1)
                val C56 = row56.createCell(2)
                val D56 = row56.createCell(3)
                val E56 = row56.createCell(4)
                val F56 = row56.createCell(5)
                B56.setCellValue("1) Depositi bancari e postali")
                disponibilitàliquide1 = findViewById(R.id.disponibilitàliquide_1)
                if (disponibilitàliquide1!!.getText().toString().length != 0) {
                    val displ1 = disponibilitàliquide1!!.getText().toString().toDouble()
                    C56.setCellValue(displ1)
                }
                //57 RIGA
                val row57 = sheet.createRow(56)
                val A57 = row57.createCell(0)
                val B57 = row57.createCell(1)
                val C57 = row57.createCell(2)
                val D57 = row57.createCell(3)
                val E57 = row57.createCell(4)
                val F57 = row57.createCell(5)
                B57.setCellValue("2) Assegni")
                D57.setCellValue("E)")
                D57.setCellStyle(grassetto)
                E57.setCellValue("RATEI E RISCONTI")
                E57.setCellStyle(grassetto)
                F57.cellFormula = "F58"
                F57.setCellStyle(grassetto)
                disponibilitàliquide2 = findViewById(R.id.disponibilitàliquide_2)
                if (disponibilitàliquide2!!.getText().toString().length != 0) {
                    val displ12 = disponibilitàliquide2!!.getText().toString().toDouble()
                    C57.setCellValue(displ12)
                }
                //58 RIGA
                val row58 = sheet.createRow(57)
                val A58 = row58.createCell(0)
                val B58 = row58.createCell(1)
                val C58 = row58.createCell(2)
                val D58 = row58.createCell(3)
                val E58 = row58.createCell(4)
                val F58 = row58.createCell(5)
                B58.setCellValue("3) Denaro e valori in cassa")
                E58.setCellValue("Ratei passivi")
                rateieriscontipassivi1 = findViewById(R.id.rateieriscontipassivi_1)
                if (rateieriscontipassivi1!!.getText().toString().length != 0) {
                    val rerp1 = rateieriscontipassivi1!!.getText().toString().toDouble()
                    F58.setCellValue(rerp1)
                }
                disponibilitàliquide3 = findViewById(R.id.disponibilitàliquide_3)
                if (disponibilitàliquide3!!.getText().toString().length != 0) {
                    val displ3 = disponibilitàliquide3!!.getText().toString().toDouble()
                    C58.setCellValue(displ3)
                }
                //59 RIGA
                val row59 = sheet.createRow(58)
                val A59 = row59.createCell(0)
                val B59 = row59.createCell(1)
                val C59 = row59.createCell(2)
                val D59 = row59.createCell(3)
                val E59 = row59.createCell(4)
                val F59 = row59.createCell(5)
                A59.setCellValue("D)")
                A59.setCellStyle(grassetto)
                B59.setCellValue("RATEI E RISCONTI")
                B59.setCellStyle(grassetto)
                C59.setCellStyle(grassetto)
                rateieriscontiattivi1 = findViewById(R.id.rateieriscontiattivi_1)
                if (rateieriscontiattivi1!!.getText().toString().length != 0) {
                    val rera1 = rateieriscontiattivi1!!.getText().toString().toDouble()
                    C59.setCellValue(rera1)
                }


                //60 RIGA
                val row60 = sheet.createRow(59)
                val A60 = row60.createCell(0)
                val B60 = row60.createCell(1)
                val C60 = row60.createCell(2)
                val D60 = row60.createCell(3)
                val E60 = row60.createCell(4)
                val F60 = row60.createCell(5)
                //61 RIGA
                val row61 = sheet.createRow(60)
                val A61 = row61.createCell(0)
                val B61 = row61.createCell(1)
                val C61 = row61.createCell(2)
                val D61 = row61.createCell(3)
                val E61 = row61.createCell(4)
                val F61 = row61.createCell(5)
                A61.setCellValue("TOTALE ATTIVO")
                A61.setCellStyle(alcentro)
                C61.cellFormula = "C5+C7+C28+C59"
                C61.setCellStyle(grassetto)
                D61.setCellValue("TOTALE PASSIVO")
                D61.setCellStyle(alcentro)
                E61.setCellStyle(style)
                F61.cellFormula = "F5+F18+F25+F28+F57"
                F61.setCellStyle(grassetto)
                RegionUtil.setBorderBottom(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A61:F61"), sheet
                )
                RegionUtil.setBorderTop(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A3:F3"), sheet
                )
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("C3:C61"), sheet
                )
                RegionUtil.setBorderRight(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("F3:F61"), sheet
                )
                RegionUtil.setBorderLeft(
                    BorderStyle.MEDIUM,
                    CellRangeAddress.valueOf("A3:A61"), sheet
                )
                RegionUtil.setBorderRight(
                    sottile.borderRight,
                    CellRangeAddress.valueOf("A4:A60"),
                    sheet
                )
                RegionUtil.setBorderRight(
                    sottile.borderRight,
                    CellRangeAddress.valueOf("B4:B61"),
                    sheet
                )
                RegionUtil.setBorderRight(
                    sottile.borderRight,
                    CellRangeAddress.valueOf("D4:D60"),
                    sheet
                )
                RegionUtil.setBorderRight(
                    sottile.borderRight,
                    CellRangeAddress.valueOf("E4:E61"),
                    sheet
                )
                for (i in 4..61) {
                    RegionUtil.setBorderTop(
                        BorderStyle.THIN, CellRangeAddress.valueOf(
                            "A$i:F$i"
                        ), sheet
                    )
                }


                // IMPOSTARE LARGHEZZA COLONNE
                sheet.setColumnWidth(1, 10000)
                sheet.setColumnWidth(4, 10000)
                val sheet2 = workbook.createSheet("SP Ricl.criterio finanziario")
                val roww1 = sheet2.createRow(0)
                val roww2 = sheet2.createRow(1)
                val roww3 = sheet2.createRow(2)
                val roww4 = sheet2.createRow(3)
                val roww5 = sheet2.createRow(4)
                val roww6 = sheet2.createRow(5)
                val roww7 = sheet2.createRow(6)
                val roww8 = sheet2.createRow(7)
                val roww9 = sheet2.createRow(8)
                val roww10 = sheet2.createRow(9)
                val roww11 = sheet2.createRow(10)
                val roww12 = sheet2.createRow(11)
                val roww13 = sheet2.createRow(12)
                val roww14 = sheet2.createRow(13)
                val roww15 = sheet2.createRow(14)
                val roww16 = sheet2.createRow(15)
                val roww17 = sheet2.createRow(16)
                val roww18 = sheet2.createRow(17)
                val roww19 = sheet2.createRow(18)
                val roww20 = sheet2.createRow(19)
                val roww21 = sheet2.createRow(20)
                val roww22 = sheet2.createRow(21)
                val roww23 = sheet2.createRow(22)
                val roww24 = sheet2.createRow(23)
                val roww25 = sheet2.createRow(24)
                val roww26 = sheet2.createRow(25)
                val roww27 = sheet2.createRow(26)
                val roww28 = sheet2.createRow(27)
                val roww29 = sheet2.createRow(28)
                val roww30 = sheet2.createRow(29)
                val roww31 = sheet2.createRow(30)
                val roww32 = sheet2.createRow(31)
                val roww33 = sheet2.createRow(32)
                val roww34 = sheet2.createRow(33)
                val roww35 = sheet2.createRow(34)
                sheet2.addMergedRegion(CellRangeAddress.valueOf("A1:D1"))
                val aa1: Cell = roww1.createCell(0)
                val bb1: Cell = roww1.createCell(1)
                val cc1: Cell = roww1.createCell(2)
                val dd1: Cell = roww1.createCell(3)
                aa1.setCellValue("STATO PATRIMONIALE RICLASSIFICATO SEGUENDO IL CRITERIO FINANZIARIO")
                //2 RIGA
                val aa2: Cell = roww2.createCell(0)
                val bb2: Cell = roww2.createCell(1)
                val cc2: Cell = roww2.createCell(2)
                val dd2: Cell = roww2.createCell(3)
                //3 RIGA
                val aa3: Cell = roww3.createCell(0)
                val bb3: Cell = roww3.createCell(1)
                val cc3: Cell = roww3.createCell(2)
                val dd3: Cell = roww3.createCell(3)
                aa3.setCellValue("ATTIVO")
                bb3.setCellValue("data")
                cc3.setCellValue("PASSIVO")
                dd3.setCellValue("data")
                //4 RIGA
                val aa4: Cell = roww4.createCell(0)
                val bb4: Cell = roww4.createCell(1)
                val cc4: Cell = roww4.createCell(2)
                val dd4: Cell = roww4.createCell(3)
                bb4.setCellValue("Val.ass.")
                dd4.setCellValue("Val.ass.")
                //5 RIGA
                val aa5: Cell = roww5.createCell(0)
                val bb5: Cell = roww5.createCell(1)
                val cc5: Cell = roww5.createCell(2)
                val dd5: Cell = roww5.createCell(3)
                //6 RIGA
                val aa6: Cell = roww6.createCell(0)
                val bb6: Cell = roww6.createCell(1)
                val cc6: Cell = roww6.createCell(2)
                val dd6: Cell = roww6.createCell(3)
                aa6.setCellValue("IMMOBILIZZAZIONI")
                bb6.cellFormula = "B7+B10+B13+B16"
                cc6.setCellValue("PATRIMONIO NETTO")
                dd6.cellFormula = "SUM(D7:D17)"
                //7 RIGA
                val aa7: Cell = roww7.createCell(0)
                val bb7: Cell = roww7.createCell(1)
                val cc7: Cell = roww7.createCell(2)
                val dd7: Cell = roww7.createCell(3)
                aa7.setCellValue("Immobilizzazioni Immateriali")
                bb7.cellFormula = "SUM(B8:B9)"
                cc7.setCellValue("I. Capitale sociale")
                if (patrimonionetto1!!.getText().toString().length != 0) {
                    val patnetto1 = patrimonionetto1!!.getText().toString().toDouble()
                    cellF6.setCellValue(patnetto1)
                    dd7.setCellValue(patnetto1)
                }

                //8 RIGA
                val aa8: Cell = roww8.createCell(0)
                val bb8: Cell = roww8.createCell(1)
                val cc8: Cell = roww8.createCell(2)
                val dd8: Cell = roww8.createCell(3)
                //9 RIGA
                val aa9: Cell = roww9.createCell(0)
                val bb9: Cell = roww9.createCell(1)
                val cc9: Cell = roww9.createCell(2)
                val dd9: Cell = roww9.createCell(3)
                //10 RIGA
                val aa10: Cell = roww10.createCell(0)
                val bb10: Cell = roww10.createCell(1)
                val cc10: Cell = roww10.createCell(2)
                val dd10: Cell = roww10.createCell(3)
                //11 RIGA
                val aa11: Cell = roww11.createCell(0)
                val bb11: Cell = roww11.createCell(1)
                val cc11: Cell = roww11.createCell(2)
                val dd11: Cell = roww11.createCell(3)
                //12 RIGA
                val aa12: Cell = roww12.createCell(0)
                val bb12: Cell = roww12.createCell(1)
                val cc12: Cell = roww12.createCell(2)
                val dd12: Cell = roww12.createCell(3)
                //13 RIGA
                val aa13: Cell = roww13.createCell(0)
                val bb13: Cell = roww13.createCell(1)
                val cc13: Cell = roww13.createCell(2)
                val dd13: Cell = roww13.createCell(3)
                //14 RIGA
                val aa14: Cell = roww14.createCell(0)
                val bb14: Cell = roww14.createCell(1)
                val cc14: Cell = roww14.createCell(2)
                val dd14: Cell = roww14.createCell(3)
                //15 RIGA
                val aa15: Cell = roww15.createCell(0)
                val bb15: Cell = roww15.createCell(1)
                val cc15: Cell = roww15.createCell(2)
                val dd15: Cell = roww15.createCell(3)
                //16 RIGA
                val aa16: Cell = roww16.createCell(0)
                val bb16: Cell = roww16.createCell(1)
                val cc16: Cell = roww16.createCell(2)
                val dd16: Cell = roww16.createCell(3)
                //17 RIGA
                val aa17: Cell = roww17.createCell(0)
                val bb17: Cell = roww17.createCell(1)
                val cc17: Cell = roww17.createCell(2)
                val dd17: Cell = roww17.createCell(3)
                //18 RIGA
                val aa18: Cell = roww18.createCell(0)
                val bb18: Cell = roww18.createCell(1)
                val cc18: Cell = roww18.createCell(2)
                val dd18: Cell = roww18.createCell(3)
                //19 RIGA
                val aa19: Cell = roww19.createCell(0)
                val bb19: Cell = roww19.createCell(1)
                val cc19: Cell = roww19.createCell(2)
                val dd19: Cell = roww19.createCell(3)
                //20 RIGA
                val aa20: Cell = roww20.createCell(0)
                val bb20: Cell = roww20.createCell(1)
                val cc20: Cell = roww20.createCell(2)
                val dd20: Cell = roww20.createCell(3)
                //21 RIGA
                val aa21: Cell = roww21.createCell(0)
                val bb21: Cell = roww21.createCell(1)
                val cc21: Cell = roww21.createCell(2)
                val dd21: Cell = roww21.createCell(3)
                //22 RIGA
                val aa22: Cell = roww22.createCell(0)
                val bb22: Cell = roww22.createCell(1)
                val cc22: Cell = roww22.createCell(2)
                val dd22: Cell = roww22.createCell(3)
                //23 RIGA
                val aa23: Cell = roww23.createCell(0)
                val bb23: Cell = roww23.createCell(1)
                val cc23: Cell = roww23.createCell(2)
                val dd23: Cell = roww23.createCell(3)
                //24 RIGA
                val aa24: Cell = roww24.createCell(0)
                val bb24: Cell = roww24.createCell(1)
                val cc24: Cell = roww24.createCell(2)
                val dd24: Cell = roww24.createCell(3)
                //25 RIGA
                val aa25: Cell = roww25.createCell(0)
                val bb25: Cell = roww25.createCell(1)
                val cc25: Cell = roww25.createCell(2)
                val dd25: Cell = roww25.createCell(3)
                //26 RIGA
                val aa26: Cell = roww26.createCell(0)
                val bb26: Cell = roww26.createCell(1)
                val cc26: Cell = roww26.createCell(2)
                val dd26: Cell = roww26.createCell(3)
                //27 RIGA
                val aa27: Cell = roww27.createCell(0)
                val bb27: Cell = roww27.createCell(1)
                val cc27: Cell = roww27.createCell(2)
                val dd27: Cell = roww27.createCell(3)
                //28 RIGA
                val aa28: Cell = roww28.createCell(0)
                val bb28: Cell = roww28.createCell(1)
                val cc28: Cell = roww28.createCell(2)
                val dd28: Cell = roww28.createCell(3)
                //29 RIGA
                val aa29: Cell = roww29.createCell(0)
                val bb29: Cell = roww29.createCell(1)
                val cc29: Cell = roww29.createCell(2)
                val dd29: Cell = roww29.createCell(3)
                //30 RIGA
                val aa30: Cell = roww30.createCell(0)
                val bb30: Cell = roww30.createCell(1)
                val cc30: Cell = roww30.createCell(2)
                val dd30: Cell = roww30.createCell(3)
                //31 RIGA
                val aa31: Cell = roww31.createCell(0)
                val bb31: Cell = roww31.createCell(1)
                val cc31: Cell = roww31.createCell(2)
                val dd31: Cell = roww31.createCell(3)
                //32 RIGA
                val aa32: Cell = roww32.createCell(0)
                val bb32: Cell = roww32.createCell(1)
                val cc32: Cell = roww32.createCell(2)
                val dd32: Cell = roww32.createCell(3)
                //33 RIGA
                val aa33: Cell = roww33.createCell(0)
                val bb33: Cell = roww33.createCell(1)
                val cc33: Cell = roww33.createCell(2)
                val dd33: Cell = roww33.createCell(3)
                //34 RIGA
                val aa34: Cell = roww34.createCell(0)
                val bb34: Cell = roww34.createCell(1)
                val cc34: Cell = roww34.createCell(2)
                val dd34: Cell = roww34.createCell(3)
                //35 RIGA
                val aa35: Cell = roww35.createCell(0)
                val bb35: Cell = roww35.createCell(1)
                val cc35: Cell = roww35.createCell(2)
                val dd35: Cell = roww35.createCell(3)
                val evaluator: FormulaEvaluator = workbook.creationHelper.createFormulaEvaluator()
                evaluator.evaluateAll()
                fileSaver.execute(workbook)
            }
        })
    }

    companion object {
        private val REQUEST_CODE_WRITE_EXTERNAL_STORAGE = 1
    }
}
