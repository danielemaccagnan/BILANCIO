package com.example.bilancio;

import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.text.Editable;
import android.text.TextWatcher;
import android.view.MenuItem;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.PopupMenu;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public class statopatrimoniale extends AppCompatActivity {
    Context context=this;
    private StatoPatrimonialeSaver fileSaverTask;
    private static final int REQUEST_CODE_WRITE_EXTERNAL_STORAGE = 1;
    // Dichiarazione del metodo showPopupMenu()
    private void showPopupMenu(View v) {
        PopupMenu popup = new PopupMenu(this, v);
        popup.show();
    }
    public EditText
    creditiversosoci1,
    immobilizzazioni1, immobilizzazioni2, immobilizzazioni3, immobilizzazioni4,immobilizzazioni5,immobilizzazioni6,immobilizzazioni7,
    immobilizzazionimateriali1,immobilizzazionimateriali2,immobilizzazionimateriali3, immobilizzazionimaterialifa, immobilizzazionimateriali4,immobilizzazionimateriali5,
    immobilizzazionifinanziarie1, immobilizzazionifinanziarie2, immobilizzazionifinanziarie3,immobilizzazionifinanziarie4,
    rimanenze1, rimanenze2, rimanenze3, rimanenze4, rimanenze5, cambialiattive,
    crediti1punto1, crediti1punto2,crediti2,crediti3,crediti4,crediti5,crediti5punto2, crediti5punto3, crediti5punto4,
    immobilizzazionifinanziarienonimmobilizzate1, immobilizzazionifinanziarienonimmobilizzate2, immobilizzazionifinanziarienonimmobilizzate3, immobilizzazionifinanziarienonimmobilizzate3punto2, immobilizzazionifinanziarienonimmobilizzate4, immobilizzazionifinanziarienonimmobilizzate5, immobilizzazionifinanziarienonimmobilizzate6, immobilizzazionifinanziarienonimmobilizzate7,
    disponibilitàliquide1, disponibilitàliquide2, disponibilitàliquide3,
    rateieriscontiattivi1,
    fatturedaemettere,
    mutuipassivi, mutuiattivi, fornitoriacconti, clientiacconti, azionistisott, erariociv, fornitoriimmobilizzazioniimmaterialicacconti, getFornitoriimmobilizzazionimaterialicacconti, imballaggidurevoli, crediticommdiversi, clienticostianticipati, cambialiallosconto, cambialiallincasso, cambialiinsolute, fondorischisucrediti, fondosvalutazionecrediti, creditiperimposte, creditiversoistitutiprevidenziali, creditipercauzioni, creditidiversi, bancaccattivi, bancaccpostali, fondoresponsabilitàcivile, fondomanutenzioniprogrammate, debitipertfr, bancacriba, bancaccpassivi, bancacanticipasufattura, devitivaltrifinanziatori, debiticommercialidiversi, debitidaliquidare, clienticacconti, cambialipassive, debitiperritenutedaversare, debitiperiva, debitiperimposte, erariociva, debitipercauzioni, personalecretribuzione, personalecliquidazione, debitivfpensione, cedenteccessione, debitidiversi,
    patrimonionetto1, patrimonionetto2, patrimonionetto3, patrimonionetto4, patrimonionetto5, patrimonionetto6, patrimonionetto7, patrimonionetto8, patrimonionetto9, patrimonionetto10,
    fondiperrischieoneri1, fondiperrischieoneri2, fondiperrischieoneri3, fondiperrischieoneri4,
    tfr1, fatturedaricevere,
    debiti1, debiti2, debiti3, debiti4, debiti4punto1, debiti4punto2, debiti5, debiti5punto1, debiti5punto2, debiti6, debiti7, debiti8, debiti9, debiti10, debiti11, debiti11punto2, debiti12, debiti13, debiti14,
    rateieriscontipassivi1;


        @Override
        protected void onCreate(Bundle savedInstanceState) {
            Context context = this;

            // inizializza l'istanza della classe ExcelFileSaver
            StatoPatrimonialeSaver fileSaver = new StatoPatrimonialeSaver(this);

            super.onCreate(savedInstanceState);
            setContentView(R.layout.activity_statopatrimoniale);



            // Trova il pulsante con id "popup_button" nella tua attività
            Button popup = findViewById(R.id.popup_button);

// Aggiungi un listener di click al pulsante
            popup.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View view) {
                    // Crea un nuovo oggetto PopupMenu e lo mostra
                    PopupMenu popupMenu = new PopupMenu(statopatrimoniale.this, view);
                    popupMenu.getMenuInflater().inflate(R.menu.popup_menu, popupMenu.getMenu());

                    // Aggiungi un listener per gestire il click degli elementi del menu
                    popupMenu.setOnMenuItemClickListener(new PopupMenu.OnMenuItemClickListener() {
                        @Override
                        public boolean onMenuItemClick(MenuItem item) {
                            switch (item.getItemId()) {
                                case R.id.menu_item_conto_economico:
                                    // Avvia la nuova activity ContoEconomicoActivity quando l'utente seleziona l'opzione Conto Economico
                                    Intent intent = new Intent(statopatrimoniale.this, contoeconomico.class);
                                    startActivity(intent);
                                    return true;
                                case R.id.menu_item_home:
                                    // Torna alla home
                                    Intent intent1 = new Intent(statopatrimoniale.this, MainActivity.class);
                                    startActivity(intent1);
                                    return true;
                                default:
                                    return false;
                            }
                        }
                    });

                    // Mostra il popup menu
                    popupMenu.show();
                }
            });



            if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED) {
                // Se il permesso non è stato ancora concesso, richiedilo all'utente
                ActivityCompat.requestPermissions(this,
                        new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        REQUEST_CODE_WRITE_EXTERNAL_STORAGE);
            }



// EditText con un solo "."
        EditText debiti11punto2 = findViewById(R.id.debiti_11_2);
        debiti11punto2.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    debiti11punto2.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    debiti11punto2.setSelection(debiti11punto2.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    debiti11punto2.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    debiti11punto2.setSelection(debiti11punto2.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });
        // EditText con un solo "."
        EditText patrimonionetto7 = findViewById(R.id.patrimonionetto_7);
        patrimonionetto7.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    patrimonionetto7.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    patrimonionetto7.setSelection(patrimonionetto7.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    patrimonionetto7.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    patrimonionetto7.setSelection(patrimonionetto7.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });
        // EditText con un solo "."
        EditText patrimonionetto10 = findViewById(R.id.patrimonionetto_10);
        patrimonionetto10.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    patrimonionetto10.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    patrimonionetto10.setSelection(patrimonionetto10.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    patrimonionetto10.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    patrimonionetto10.setSelection(patrimonionetto10.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });
        // EditText con un solo "."
        EditText immobilizzazionifinanziarienonimmobilizzate3punto2 = findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_3_2);
        immobilizzazionifinanziarienonimmobilizzate3punto2.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setSelection(immobilizzazionifinanziarienonimmobilizzate3punto2.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    immobilizzazionifinanziarienonimmobilizzate3punto2.setSelection(immobilizzazionifinanziarienonimmobilizzate3punto2.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });
        //Edit text con un solo .
        EditText crediti5 = findViewById(R.id.crediti_5);
        crediti5.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    crediti5.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    crediti5.setSelection(crediti5.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    crediti5.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    crediti5.setSelection(crediti5.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });

// EditText con solo numeri e un solo "."
        EditText debiti13 = findViewById(R.id.debiti_13);
        debiti13.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence charSequence, int i, int i1, int i2) {
            }

            @Override
            public void onTextChanged(CharSequence charSequence, int i, int i1, int i2) {
                // Controlla se viene inserito un secondo "."
                if (charSequence.toString().contains(".") && charSequence.toString().lastIndexOf(".") != charSequence.toString().indexOf(".")) {
                    // Rimuovi il secondo "."
                    debiti13.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf(".")) + charSequence.toString().substring(charSequence.toString().lastIndexOf(".") + 1));
                    debiti13.setSelection(debiti13.getText().length());
                }
                if (charSequence.toString().contains("-") && charSequence.toString().lastIndexOf("-") != charSequence.toString().indexOf("-")) {
                    // Rimuovi il secondo "."
                    debiti13.setText(charSequence.toString().substring(0, charSequence.toString().lastIndexOf("-")) + charSequence.toString().substring(charSequence.toString().lastIndexOf("-") + 1));
                    debiti13.setSelection(debiti13.getText().length());
                }
            }

            @Override
            public void afterTextChanged(Editable editable) {
            }
        });


        Button salvaBilancio = findViewById(R.id.salva);



        salvaBilancio.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                // Creazione del workbook
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Stato Patrimoniale");

                String fileName = "statopatrimoniale.xlsx";
                int counter = 1;
                File file = new File(context.getExternalFilesDir(null), fileName);
                while (file.exists()) {
                    fileName = "statopatrimoniale(" + counter + ").xlsx";
                    file = new File(context.getExternalFilesDir(null), fileName);
                    counter++;
                }

                // crea una nuova istanza di ExcelFileSaver e esegui il salvataggio
                StatoPatrimonialeSaver fileSaver = new StatoPatrimonialeSaver(context);






                XSSFRow row = sheet.createRow(0);

                sheet.addMergedRegion(CellRangeAddress.valueOf("A1:F1"));

                // Creare stile di cella
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                // Crea un nuovo font e imposta la proprietà "bold" su true
                Font font = workbook.createFont();
                font.setBold(true);

                // Crea un nuovo stile di cella e imposta il font creato sopra
                CellStyle grassetto = workbook.createCellStyle();
                grassetto.setFont(font);
             CellStyle alcentro = workbook.createCellStyle();
             alcentro.setAlignment(HorizontalAlignment.CENTER);
             alcentro.setFont(font);
                // Creare un nuovo oggetto CellStyle e impostare l'allineamento orizzontale al centro
                CellStyle stile = workbook.createCellStyle();
                stile.setAlignment(HorizontalAlignment.CENTER);
                stile.setFont(font);

                CellStyle sottile = workbook.createCellStyle();
                sottile.setBorderRight(BorderStyle.THIN);
                sottile.setBorderTop(BorderStyle.THIN);





             // Creare una cella e scrivere il testo "Stato Patrimoniale" al suo interno
                Cell statopatrimoniale = row.createCell(0);
                statopatrimoniale.setCellValue("STATO PATRIMONIALE");
                statopatrimoniale.setCellStyle(alcentro);
                // 2 RIGA
             XSSFRow row2 = sheet.createRow(1);
             XSSFCell A2 = row2.createCell(0);
             XSSFCell B2 = row2.createCell(1);
             XSSFCell C2 = row2.createCell(2);
             XSSFCell D2 = row2.createCell(3);
             XSSFCell E2 = row2.createCell(4);
             XSSFCell F2 = row2.createCell(5);

             sheet.addMergedRegion(CellRangeAddress.valueOf("A3:B3"));
             sheet.addMergedRegion(CellRangeAddress.valueOf("D3:E3"));
             sheet.addMergedRegion(CellRangeAddress.valueOf("A61:B61"));
             sheet.addMergedRegion(CellRangeAddress.valueOf("D61:E61"));
             // 3 RIGA

                XSSFRow row3 = sheet.createRow(2);
                XSSFCell cellA3 = row3.createCell(0);
                XSSFCell cellB3 = row3.createCell(1);
                XSSFCell cellC3 = row3.createCell(2);
                XSSFCell cellD3 = row3.createCell(3);
                XSSFCell cellE3 = row3.createCell(4);
                XSSFCell cellF3 = row3.createCell(5);
                cellA3.setCellValue("ATTIVO");
                cellA3.setCellStyle(alcentro);
                cellC3.setCellValue("data");
                cellC3.setCellStyle(grassetto);
                cellD3.setCellValue("PASSIVO");
                cellD3.setCellStyle(alcentro);
                cellF3.setCellValue("data");
                cellF3.setCellStyle(grassetto);



                //5 RIGA
                XSSFRow row5 = sheet.createRow(4);
                XSSFCell cellA5 = row5.createCell(0);
                XSSFCell cellB5 = row5.createCell(1);
                XSSFCell cellC5 = row5.createCell(2);
                XSSFCell cellD5 = row5.createCell(3);
                XSSFCell cellE5 = row5.createCell(4);
                XSSFCell cellF5 = row5.createCell(5);
               cellA5.setCellValue("A)");
               cellA5.setCellStyle(grassetto);
               cellB5.setCellValue("CREDITI VERSO SOCI PER VERSAMENTI ANCORA DOVUTI");
               cellB5.setCellStyle(grassetto);
               cellC5.setCellStyle(grassetto);
               cellD5.setCellValue("A)");
               cellD5.setCellStyle(grassetto);
               cellE5.setCellValue("PATRIMONIO NETTO");
               cellE5.setCellStyle(grassetto);
               cellF5.setCellFormula("SUM(F6:F16)");
               cellF5.setCellStyle(grassetto);
                creditiversosoci1=findViewById(R.id.creditiversosoci_1);
                azionistisott = findViewById(R.id.azionisticsott);
                if(creditiversosoci1.getText().toString().length()!=0||azionistisott.getText().toString().length()!=0){
                    double crevssoci1 = creditiversosoci1.getText().toString().trim().isEmpty()?0: Double.parseDouble(creditiversosoci1.getText().toString());
                    double azionistisotto = azionistisott.getText().toString().trim().isEmpty()?0: Double.parseDouble(azionistisott.getText().toString());
                    cellC5.setCellValue(crevssoci1 + azionistisotto);
                }


                // SESTA RIGA
                XSSFRow row6 = sheet.createRow(5);
                XSSFCell cellE6 = row6.createCell(4);
                XSSFCell cellF6 = row6.createCell(5);
                cellE6.setCellValue("I. Capitale sociale");
                patrimonionetto1=findViewById(R.id.patrimonionetto_1);



                // SETTIMA RIGA
                XSSFRow row7 = sheet.createRow(6);
                XSSFCell cellA7 = row7.createCell(0);
                XSSFCell cellB7 = row7.createCell(1);
                XSSFCell cellC7 = row7.createCell(2);
                XSSFCell cellE7 = row7.createCell(4);
                XSSFCell cellF7 = row7.createCell(5);
                patrimonionetto2=findViewById(R.id.patrimonionetto_2);
                if(patrimonionetto2.getText().toString().length()!=0){
                    double patnetto2=Double.parseDouble(patrimonionetto2.getText().toString());
                    cellF7.setCellValue(patnetto2);
                }

                cellA7.setCellValue("B)");
                cellA7.setCellStyle(grassetto);
                cellB7.setCellValue("IMMOBILIZZAZIONI");
                cellB7.setCellStyle(grassetto);
                cellC7.setCellFormula("C8+C16+C23");
                cellC7.setCellStyle(grassetto);
                if(cellC7.getNumericCellValue()==0){
                    cellC7.setCellValue("1");
                }
                cellE7.setCellValue("II. Riserva sovrapprezzo delle azioni");
                // 8 RIGA
                XSSFRow row8 = sheet.createRow(7);
                XSSFCell A8 = row8.createCell(0);
                XSSFCell B8 = row8.createCell(1);
                XSSFCell C8 = row8.createCell(2);
                XSSFCell D8 = row8.createCell(3);
                XSSFCell E8 = row8.createCell(4);
                XSSFCell F8 = row8.createCell(5);
                patrimonionetto3=findViewById(R.id.patrimonionetto_3);
                if(patrimonionetto3.getText().toString().length()!=0){
                    double patnetto3=Double.parseDouble(patrimonionetto3.getText().toString());
                    F8.setCellValue(patnetto3);
                }

                A8.setCellValue("I)");
                A8.setCellStyle(grassetto);
                B8.setCellValue("Immobilizzazioni Immateriali");
                B8.setCellStyle(grassetto);
                C8.setCellFormula("SUM(C9:C15)");
                C8.setCellStyle(grassetto);
                E8.setCellValue("III. Riserva di rivalutazione");

                //9 RIGA
                XSSFRow row9 = sheet.createRow(8);
                XSSFCell A9 = row9.createCell(0);
                XSSFCell B9 = row9.createCell(1);
                XSSFCell C9 = row9.createCell(2);
                XSSFCell D9 = row9.createCell(3);
                XSSFCell E9 = row9.createCell(4);
                XSSFCell F9 = row9.createCell(5);
                B9.setCellValue("1) Costi di impianto e ampliamento");
                E9.setCellValue("IV. Riserva legale");

                patrimonionetto4=findViewById(R.id.patrimonionetto_4);
                if(patrimonionetto4.getText().toString().length()!=0){
                    double patnetto4=Double.parseDouble(patrimonionetto4.getText().toString());
                    F9.setCellValue(patnetto4);
                }
                immobilizzazioni1=findViewById(R.id.immobilizzazioni_1);

                if(immobilizzazioni1.getText().toString().length()!=0){
                    double imm1 = Double.parseDouble(immobilizzazioni1.getText().toString());
                    C9.setCellValue(imm1);
                }



                // 10 RIGA
                XSSFRow row10 = sheet.createRow(9);
                XSSFCell A10 = row10.createCell(0);
                XSSFCell B10 = row10.createCell(1);
                XSSFCell C10 = row10.createCell(2);
                XSSFCell D10 = row10.createCell(3);
                XSSFCell E10 = row10.createCell(4);
                XSSFCell F10 = row10.createCell(5);
                immobilizzazioni2=findViewById(R.id.immobilizzazioni_2);
                if(immobilizzazioni2.getText().toString().length()!=0){
                    double imm2 = Double.parseDouble(immobilizzazioni2.getText().toString());
                    C10.setCellValue(imm2);
                }

                patrimonionetto5=findViewById(R.id.patrimonionetto_5);
                if(patrimonionetto5.getText().toString().length()!=0){
                    double patnetto5=Double.parseDouble(patrimonionetto5.getText().toString());
                    F10.setCellValue(patnetto5);
                }
                B10.setCellValue("2) Costi  di sviluppo");
                B10.setCellStyle(grassetto);
                E10.setCellValue("V. Riserve statutarie");
                // 11 RIGA
                XSSFRow row11 = sheet.createRow(10);
                XSSFCell A11 = row11.createCell(0);
                XSSFCell B11 = row11.createCell(1);
                XSSFCell C11 = row11.createCell(2);
                XSSFCell D11 = row11.createCell(3);
                XSSFCell E11 = row11.createCell(4);
                XSSFCell F11 = row11.createCell(5);

                patrimonionetto6=findViewById(R.id.patrimonionetto_6);
                if(patrimonionetto6.getText().toString().length()!=0){
                    double patnetto6=Double.parseDouble(patrimonionetto6.getText().toString());
                    F11.setCellValue(patnetto6);
                }


                immobilizzazioni3=findViewById(R.id.immobilizzazioni_3);

                if(immobilizzazioni3.getText().toString().length()!=0){
                    double imm3 = Double.parseDouble(immobilizzazioni3.getText().toString());
                    C11.setCellValue(imm3);
                }

                B11.setCellValue("3) Diritti di brevetto industriale");
                E11.setCellValue("VI. Altre riserve, distintamente indicate");
                // 12 RIGA
                XSSFRow row12 = sheet.createRow(11);
                XSSFCell A12 = row12.createCell(0);
                XSSFCell B12 = row12.createCell(1);
                XSSFCell C12 = row12.createCell(2);
                XSSFCell D12 = row12.createCell(3);
                XSSFCell E12 = row12.createCell(4);
                XSSFCell F12 = row12.createCell(5);

               // patrimonionetto7=findViewById(R.id.patrimonionetto_7);
                if(patrimonionetto7.getText().toString().length()!=0){
                    double patnetto7=Double.parseDouble(patrimonionetto7.getText().toString());
                    F12.setCellValue(patnetto7);
                }

                immobilizzazioni4=findViewById(R.id.immobilizzazioni_4);

                if(immobilizzazioni4.getText().toString().length()!=0){
                    double imm4 = Double.parseDouble(immobilizzazioni4.getText().toString());
                    C12.setCellValue(imm4);
                }
                B12.setCellValue("4) Concessioni, licenze,  marchi e diritti simili");
                E12.setCellValue("VII. Riserva per operazioni di copertura dei flussi finanziari attesi");
                E12.setCellStyle(grassetto);
                // 13 RIGA
                XSSFRow row13 = sheet.createRow(12);
                XSSFCell A13 = row13.createCell(0);
                XSSFCell B13 = row13.createCell(1);
                XSSFCell C13 = row13.createCell(2);
                XSSFCell D13 = row13.createCell(3);
                XSSFCell E13 = row13.createCell(4);
                XSSFCell F13 = row13.createCell(5);

                patrimonionetto8=findViewById(R.id.patrimonionetto_8);
                if(patrimonionetto8.getText().toString().length()!=0){
                    double patnetto8=Double.parseDouble(patrimonionetto8.getText().toString());
                    F13.setCellValue(patnetto8);
                }

                immobilizzazioni5=findViewById(R.id.immobilizzazioni_5);

                if(immobilizzazioni5.getText().toString().length()!=0){
                    double imm5 = Double.parseDouble(immobilizzazioni5.getText().toString());
                    C13.setCellValue(imm5);
                }
                B13.setCellValue("5) Avviamento");
                E13.setCellValue("VIII. Utili (perdite) portati a nuovo");
                // 14 RIGA
                XSSFRow row14 = sheet.createRow(13);
                XSSFCell A14 = row14.createCell(0);
                XSSFCell B14 = row14.createCell(1);
                XSSFCell C14 = row14.createCell(2);
                XSSFCell D14 = row14.createCell(3);
                XSSFCell E14 = row14.createCell(4);
                XSSFCell F14 = row14.createCell(5);

                patrimonionetto9=findViewById(R.id.patrimonionetto_9);
                if(patrimonionetto9.getText().toString().length()!=0){
                    double patnetto9=Double.parseDouble(patrimonionetto9.getText().toString());
                    F14.setCellValue(patnetto9);
                }

                immobilizzazioni6=findViewById(R.id.immobilizzazioni_6);
                fornitoriimmobilizzazioniimmaterialicacconti = findViewById(R.id.fornitori_imm_imm_cacconti);
                if(immobilizzazioni6.getText().toString().length()!=0||fornitoriimmobilizzazioniimmaterialicacconti.getText().toString().length()!=0){
                    double imm6 =immobilizzazioni6.getText().toString().trim().isEmpty()?0: Double.parseDouble(immobilizzazioni6.getText().toString());
                    double fornitoriimmcacconti = fornitoriimmobilizzazioniimmaterialicacconti.getText().toString().trim().isEmpty()?0: Double.parseDouble(fornitoriimmobilizzazioniimmaterialicacconti.getText().toString());
                    C14.setCellValue(imm6 + fornitoriimmcacconti);
                }
                B14.setCellValue("6) Immobilizzazioni in corso e acconti");
                E14.setCellValue("IX. Utile (perdita) d'esercizio");

                // 15 RIGA
                XSSFRow row15 = sheet.createRow(14);
                XSSFCell A15 = row15.createCell(0);
                XSSFCell B15 = row15.createCell(1);
                XSSFCell C15 = row15.createCell(2);
                XSSFCell D15 = row15.createCell(3);
                XSSFCell E15 = row15.createCell(4);
                XSSFCell F15 = row15.createCell(5);

               // patrimonionetto10=findViewById(R.id.patrimonionetto_10);
                if(patrimonionetto10.getText().toString().length()!=0){
                    double patnetto10=Double.parseDouble(patrimonionetto10.getText().toString());
                    F15.setCellValue(patnetto10);
                }

                immobilizzazioni7=findViewById(R.id.immobilizzazioni_7);

                if(immobilizzazioni7.getText().toString().length()!=0){
                    double imm7 = Double.parseDouble(immobilizzazioni7.getText().toString());
                    C15.setCellValue(imm7);
                }
                B15.setCellValue("7) Altre");
                E15.setCellValue("X. Riserva negativa per azioni proprie in portafoglio");
                E15.setCellStyle(grassetto);
                // 16 RIGA
                XSSFRow row16 = sheet.createRow(15);
                XSSFCell A16 = row16.createCell(0);
                XSSFCell B16 = row16.createCell(1);
                XSSFCell C16 = row16.createCell(2);
                XSSFCell D16 = row16.createCell(3);
                XSSFCell E16 = row16.createCell(4);
                XSSFCell F16 = row16.createCell(5);
                A16.setCellValue("II)");
                A16.setCellStyle(grassetto);
                B16.setCellValue("Immobilizzazioni Materiali");
                B16.setCellStyle(grassetto);
                C16.setCellFormula("SUM(C17:C22)");
                C16.setCellStyle(grassetto);
                // 17 RIGA
                XSSFRow row17 = sheet.createRow(16);
                XSSFCell A17 = row17.createCell(0);
                XSSFCell B17 = row17.createCell(1);
                XSSFCell C17 = row17.createCell(2);
                XSSFCell D17 = row17.createCell(3);
                XSSFCell E17 = row17.createCell(4);
                XSSFCell F17 = row17.createCell(5);

                immobilizzazionimateriali1=findViewById(R.id.immobilizzazionimateriali_1);
                if(immobilizzazionimateriali1.getText().toString().length()!=0){
                    double immM1 = Double.parseDouble(immobilizzazionimateriali1.getText().toString());
                    C17.setCellValue(immM1);
                }

                B17.setCellValue("1) Terreni e fabbricati");

                //18 RIGA
                XSSFRow row18 = sheet.createRow(17);
                XSSFCell A18 = row18.createCell(0);
                XSSFCell B18 = row18.createCell(1);
                XSSFCell C18 = row18.createCell(2);
                XSSFCell D18 = row18.createCell(3);
                XSSFCell E18 = row18.createCell(4);
                XSSFCell F18 = row18.createCell(5);

                immobilizzazionimateriali2=findViewById(R.id.immobilizzazionimateriali_2);
                if(immobilizzazionimateriali2.getText().toString().length()!=0){

                    double immM2 = Double.parseDouble(immobilizzazionimateriali2.getText().toString());
                    C18.setCellValue(immM2);
                }
                B18.setCellValue("2) Impianti e macchinari");
                D18.setCellValue("B)");
                D18.setCellStyle(grassetto);
                E18.setCellValue("FONDI PER RISCHI E ONERI");
                E18.setCellStyle(grassetto);
                F18.setCellFormula("SUM(F19:F24)");
                F18.setCellStyle(grassetto);
                // 19 RIGA
                XSSFRow row19 = sheet.createRow(18);
                XSSFCell A19 = row19.createCell(0);
                XSSFCell B19 = row19.createCell(1);
                XSSFCell C19 = row19.createCell(2);
                XSSFCell D19 = row19.createCell(3);
                XSSFCell E19 = row19.createCell(4);
                XSSFCell F19 = row19.createCell(5);

                immobilizzazionimateriali3=findViewById(R.id.immobilizzazionimateriali_3);
                if(immobilizzazionimateriali3.getText().toString().length()!=0){
                    double immM3 = Double.parseDouble(immobilizzazionimateriali3.getText().toString());
                    C19.setCellValue(immM3);
                }

                fondiperrischieoneri1=findViewById(R.id.fondiperrischieoneri_1);
                if(fondiperrischieoneri1.getText().toString().length()!=0){
                    double fpr1=Double.parseDouble(fondiperrischieoneri1.getText().toString());
                    F19.setCellValue(fpr1);
                }
                B19.setCellValue("3) Attrezzature industriali e commerciali");
                E19.setCellValue("1) Per trattamento di quiescenza e obblighi simili");
                // 20 RIGA
                XSSFRow row20 = sheet.createRow(19);
                XSSFCell A20 = row20.createCell(0);
                XSSFCell B20 = row20.createCell(1);
                XSSFCell C20 = row20.createCell(2);
                XSSFCell D20 = row20.createCell(3);
                XSSFCell E20 = row20.createCell(4);
                XSSFCell F20 = row20.createCell(5);

                immobilizzazionimaterialifa=findViewById(R.id.immobilizzazionimateriali_F_A);
                if(immobilizzazionimaterialifa.getText().toString().length()!=0){
                    double immMfa = Double.parseDouble(immobilizzazionimaterialifa.getText().toString());
                    C20.setCellValue(immMfa);
                }

                B20.setCellValue("F.Amm. Edifici industriali");
                E20.setCellValue("2) Per imposte");


                fondiperrischieoneri2=findViewById(R.id.fondiperrischieoneri_2);
                if(fondiperrischieoneri2.getText().toString().length()!=0){
                    double fpr2=Double.parseDouble(fondiperrischieoneri2.getText().toString());
                F20.setCellValue(fpr2);
                }

                //21 RIGA
                XSSFRow row21 = sheet.createRow(20);
                XSSFCell A21 = row21.createCell(0);
                XSSFCell B21 = row21.createCell(1);
                XSSFCell C21 = row21.createCell(2);
                XSSFCell D21 = row21.createCell(3);
                XSSFCell E21 = row21.createCell(4);
                XSSFCell F21 = row21.createCell(5);

                immobilizzazionimateriali4=findViewById(R.id.immobilizzazionimateriali_4);
                imballaggidurevoli = findViewById(R.id.imballaggidurevoli);
                if(immobilizzazionimateriali4.getText().toString().length()!=0||imballaggidurevoli.getText().toString().length()!=0){
                    double immM4 =immobilizzazionimateriali4.getText().toString().trim().isEmpty()?0: Double.parseDouble(immobilizzazionimateriali4.getText().toString());
                    double imballaggi = imballaggidurevoli.getText().toString().trim().isEmpty()?0: Double.parseDouble(imballaggidurevoli.getText().toString());
                    C21.setCellValue(immM4 + imballaggi);
                }
                B21.setCellValue("4)Altri beni");
                E21.setCellValue("3) strumenti finanziari derivati passivi");
                E21.setCellStyle(grassetto);


                fondiperrischieoneri3=findViewById(R.id.fondiperrischieoneri_3);
                if(fondiperrischieoneri3.getText().toString().length()!=0){
                    double fpr3=Double.parseDouble(fondiperrischieoneri3.getText().toString());
                    F21.setCellValue(fpr3);

                }

                // 22 RIGA
                XSSFRow row22 = sheet.createRow(21);
                XSSFCell A22 = row22.createCell(0);
                XSSFCell B22 = row22.createCell(1);
                XSSFCell C22 = row22.createCell(2);
                XSSFCell D22 = row22.createCell(3);
                XSSFCell E22 = row22.createCell(4);
                XSSFCell F22 = row22.createCell(5);


                immobilizzazionimateriali5=findViewById(R.id.immobilizzazionimateriali_5);
                fornitoriimmobilizzazioniimmaterialicacconti = findViewById(R.id.fornitori_imm_mat_cacconti);
                if(immobilizzazionimateriali5.getText().toString().length()!=0||fornitoriimmobilizzazioniimmaterialicacconti.getText().toString().length()!=0){
                    double immM5 =immobilizzazionimateriali5.getText().toString().trim().isEmpty()?0: Double.parseDouble(immobilizzazionimateriali5.getText().toString());
                    double immobilizzazionimat = fornitoriimmobilizzazioniimmaterialicacconti.getText().toString().trim().isEmpty()? 0: Double.parseDouble(fornitoriimmobilizzazioniimmaterialicacconti.getText().toString());
                    C22.setCellValue(immM5 + immobilizzazionimat);
                }

                fondiperrischieoneri4=findViewById(R.id.fondiperrischieoneri_4);
                if(fondiperrischieoneri4.getText().toString().length()!=0){
                    double fpr4=Double.parseDouble(fondiperrischieoneri4.getText().toString());
                    F22.setCellValue(fpr4);
                }
                B22.setCellValue("5) Immobilizzazioni in corso e acconti");
                E22.setCellValue("4) Altri");
                // 23 RIGA
                XSSFRow row23 = sheet.createRow(22);
                XSSFCell A23 = row23.createCell(0);
                XSSFCell B23 = row23.createCell(1);
                XSSFCell C23 = row23.createCell(2);
                XSSFCell D23 = row23.createCell(3);
                XSSFCell E23 = row23.createCell(4);
                XSSFCell F23 = row23.createCell(5);
                A23.setCellValue("III)");
                A23.setCellStyle(grassetto);
                B23.setCellValue("Immobilizzazioni Finanziarie");
                B23.setCellStyle(grassetto);
                C23.setCellFormula("SUM(C24:C27)");
                C23.setCellStyle(grassetto);

                // 24 RIGA
                XSSFRow row24 = sheet.createRow(23);
                XSSFCell A24 = row24.createCell(0);
                XSSFCell B24 = row24.createCell(1);
                XSSFCell C24 = row24.createCell(2);
                XSSFCell D24 = row24.createCell(3);
                XSSFCell E24 = row24.createCell(4);
                XSSFCell F24 = row24.createCell(5);
                B24.setCellValue("1) Partecipazioni");

                immobilizzazionifinanziarie1=findViewById(R.id.immobilizzazionifinanziarie_1);
                if(immobilizzazionifinanziarie1.getText().toString().length()!=0){
                    double immf1 = Double.parseDouble(immobilizzazionifinanziarie1.getText().toString());
                    C24.setCellValue(immf1);
                }


                // 25 RIGA
                XSSFRow row25 = sheet.createRow(24);
                XSSFCell A25 = row25.createCell(0);
                XSSFCell B25 = row25.createCell(1);
                XSSFCell C25 = row25.createCell(2);
                XSSFCell D25 = row25.createCell(3);
                XSSFCell E25 = row25.createCell(4);
                XSSFCell F25 = row25.createCell(5);
                B25.setCellValue("2) Crediti");
                D25.setCellValue("C)");
                D25.setCellStyle(grassetto);
                E25.setCellValue("TFR");
                E25.setCellStyle(grassetto);
                F25.setCellStyle(grassetto);

                mutuiattivi = findViewById(R.id.mutui_attivi);
                immobilizzazionifinanziarie2=findViewById(R.id.immobilizzazionifinanziarie_2);
                if(immobilizzazionifinanziarie2.getText().toString().length()!=0 || mutuiattivi.getText().toString().length()!=0){
                    double immf2 = immobilizzazionifinanziarie2.getText().toString().trim().isEmpty() ? 0 : Double.parseDouble(immobilizzazionifinanziarie2.getText().toString());
                    double mutuiattivii = mutuiattivi.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(mutuiattivi.getText().toString());
                    C25.setCellValue(immf2 + mutuiattivii);
                }


                tfr1=findViewById(R.id.tfr_1);
                if(tfr1.getText().toString().length()!=0){
                    double tfr=Double.parseDouble(tfr1.getText().toString());
                    F25.setCellValue(tfr);
                }

                // 26 RIGA
                XSSFRow row26 = sheet.createRow(25);
                XSSFCell A26 = row26.createCell(0);
                XSSFCell B26 = row26.createCell(1);
                XSSFCell C26 = row26.createCell(2);
                XSSFCell D26 = row26.createCell(3);
                XSSFCell E26 = row26.createCell(4);
                XSSFCell F26 = row26.createCell(5);
                B26.setCellValue("3) Altri Titoli ");

                immobilizzazionifinanziarie3=findViewById(R.id.immobilizzazionifinanziarie_3);
                if(immobilizzazionifinanziarie3.getText().toString().length()!=0){
                    double immf3=Double.parseDouble(immobilizzazionifinanziarie3.getText().toString());
                    C26.setCellValue(immf3);
                }

                // 27 RIGA
                XSSFRow row27 = sheet.createRow(26);
                XSSFCell A27 = row27.createCell(0);
                XSSFCell B27 = row27.createCell(1);
                XSSFCell C27 = row27.createCell(2);
                XSSFCell D27 = row27.createCell(3);
                XSSFCell E27 = row27.createCell(4);
                XSSFCell F27 = row27.createCell(5);
                B27.setCellValue("4) Strumenti finanziari derivati attivi");
                B27.setCellStyle(grassetto);


                immobilizzazionifinanziarie4=findViewById(R.id.immobilizzazionifinanziarie_4);
                if(immobilizzazionifinanziarie4.getText().toString().length()!=0){
                    double immf4=Double.parseDouble(immobilizzazionifinanziarie4.getText().toString());
                    C27.setCellValue(immf4);
                }

                // 28 RIGA
                XSSFRow row28 = sheet.createRow(27);
                XSSFCell A28 = row28.createCell(0);
                XSSFCell B28 = row28.createCell(1);
                XSSFCell C28 = row28.createCell(2);
                XSSFCell D28 = row28.createCell(3);
                XSSFCell E28 = row28.createCell(4);
                XSSFCell F28 = row28.createCell(5);
                A28.setCellValue("C)");
                A28.setCellStyle(grassetto);
                B28.setCellValue("ATTIVO CIRCOLANTE");
                B28.setCellStyle(grassetto);
                C28.setCellFormula("C29+C35+C55+C46");
                C28.setCellStyle(grassetto);
                D28.setCellValue("D)");
                D28.setCellStyle(grassetto);
                E28.setCellValue("DEBITI");
                E28.setCellStyle(grassetto);
                F28.setCellFormula("F29+F30+F31+F32+F35+F38+F39+F40+F41+F42+F43+F44+F45+F46+F47");
                F28.setCellStyle(grassetto);
                // 29 RIGA
                XSSFRow row29 = sheet.createRow(28);
                XSSFCell A29 = row29.createCell(0);
                XSSFCell B29 = row29.createCell(1);
                XSSFCell C29 = row29.createCell(2);
                XSSFCell D29 = row29.createCell(3);
                XSSFCell E29 = row29.createCell(4);
                XSSFCell F29 = row29.createCell(5);
                A29.setCellValue("I)");
                A29.setCellStyle(grassetto);
                B29.setCellValue("Rimanenze");
                B29.setCellStyle(grassetto);
                C29.setCellFormula("SUM(C30:C34)");
                C29.setCellStyle(grassetto);
                E29.setCellValue("1) Obbligazioni ");

                debiti1=findViewById(R.id.debiti_1);
                if(debiti1.getText().toString().length()!=0){
                    double debt1=Double.parseDouble(debiti1.getText().toString());
                    F29.setCellValue(debt1);
                }

                // 30 RIGA
                XSSFRow row30 = sheet.createRow(29);
                XSSFCell A30 = row30.createCell(0);
                XSSFCell B30 = row30.createCell(1);
                XSSFCell C30 = row30.createCell(2);
                XSSFCell D30 = row30.createCell(3);
                XSSFCell E30 = row30.createCell(4);
                XSSFCell F30 = row30.createCell(5);
                B30.setCellValue("1) Materie prime, sussidiarie e di consumo");
                E30.setCellValue("2) Obbligazioni convertibili");

                rimanenze1=findViewById(R.id.rimanenze_1);
                if(rimanenze1.getText().toString().length()!=0){
                    double rim1=Double.parseDouble(rimanenze1.getText().toString());
                    C30.setCellValue(rim1);
                }

                debiti2=findViewById(R.id.debiti_2);
                if(debiti2.getText().toString().length()!=0){
                    double debt2=Double.parseDouble(debiti2.getText().toString());
                    F30.setCellValue(debt2);
                }

                // 31 RIGA
                XSSFRow row31 = sheet.createRow(30);
                XSSFCell A31 = row31.createCell(0);
                XSSFCell B31 = row31.createCell(1);
                XSSFCell C31 = row31.createCell(2);
                XSSFCell D31 = row31.createCell(3);
                XSSFCell E31 = row31.createCell(4);
                XSSFCell F31 = row31.createCell(5);
                B31.setCellValue("2) Prodotti in corso di lavorazione e semilavorati");
                E31.setCellValue("3) Debiti verso soci per finanziamenti");

                rimanenze2=findViewById(R.id.rimanenze_2);
                if(rimanenze2.getText().toString().length()!=0){
                    double rim2=Double.parseDouble(rimanenze2.getText().toString());
                    C31.setCellValue(rim2);
                }

                debiti3=findViewById(R.id.debiti_3);
                if(debiti3.getText().toString().length()!=0){
                    double debt3=Double.parseDouble(debiti3.getText().toString());
                    F31.setCellValue(debt3);
                }

                // 32 RIGA
                XSSFRow row32 = sheet.createRow(31);
                XSSFCell A32 = row32.createCell(0);
                XSSFCell B32 = row32.createCell(1);
                XSSFCell C32 = row32.createCell(2);
                XSSFCell D32 = row32.createCell(3);
                XSSFCell E32 = row32.createCell(4);
                XSSFCell F32 = row32.createCell(5);
                B32.setCellValue("3) Lavori in corso su ordinazione");
                E32.setCellValue("4) Debiti verso banche:");
                F32.setCellFormula("F33+F34");


                rimanenze3=findViewById(R.id.rimanenze_3);
                if(rimanenze3.getText().toString().length()!=0){
                    double rim3=Double.parseDouble(rimanenze3.getText().toString());
                    C32.setCellValue(rim3);
                }
                // 33 RIGA
                XSSFRow row33 = sheet.createRow(32);
                XSSFCell A33 = row33.createCell(0);
                XSSFCell B33 = row33.createCell(1);
                XSSFCell C33 = row33.createCell(2);
                XSSFCell D33 = row33.createCell(3);
                XSSFCell E33 = row33.createCell(4);
                XSSFCell F33 = row33.createCell(5);
                B33.setCellValue("4) Prodotti finiti");
                E33.setCellValue("      - entro l'anno");

                rimanenze4=findViewById(R.id.rimanenze_4);
                if(rimanenze4.getText().toString().length()!=0){
                    double rim4=Double.parseDouble(rimanenze4.getText().toString());
                    C33.setCellValue(rim4);
                }

                debiti4punto1=findViewById(R.id.debiti_4_1);
                mutuipassivi = findViewById(R.id.mutui_passivi);
                String mutuipassiviText = mutuipassivi.getText().toString().trim();
                String debiti4punto1text = debiti4punto1.getText().toString().trim();
                if(!mutuipassiviText.isEmpty()||!debiti4punto1text.isEmpty()){
                    try{
                        double mutuipassivii = mutuipassiviText.isEmpty()? 0 : Double.parseDouble(mutuipassiviText);
                        double debt4punto1= debiti4punto1text.isEmpty()? 0: Double.parseDouble(debiti4punto1text);
                        F33.setCellValue(debt4punto1 + mutuipassivii);
                    } catch (Exception e){

                    }

                }
                // 34 RIGA
                XSSFRow row34 = sheet.createRow(33);
                XSSFCell A34 = row34.createCell(0);
                XSSFCell B34 = row34.createCell(1);
                XSSFCell C34 = row34.createCell(2);
                XSSFCell D34 = row34.createCell(3);
                XSSFCell E34 = row34.createCell(4);
                XSSFCell F34 = row34.createCell(5);
                B34.setCellValue("5) Acconti");
                E34.setCellValue("      - oltre l'anno");

                rimanenze5=findViewById(R.id.rimanenze_5);
                fornitoriacconti = findViewById(R.id.fornitori_cacconti);
                if(rimanenze5.getText().toString().length()!=0||fornitoriacconti.toString().length()!=0){
                    double rim5=rimanenze5.getText().toString().trim().isEmpty()? 0: Double.parseDouble(rimanenze5.getText().toString());
                    double fornitoricacc = fornitoriacconti.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(fornitoriacconti.getText().toString());
                    C34.setCellValue(rim5 + fornitoricacc);
                }

                debiti4punto2=findViewById(R.id.debiti_4_2);
                if(debiti4punto2.getText().toString().length()!=0){
                    double debt4punto2=Double.parseDouble(debiti4punto2.getText().toString());
                    F34.setCellValue(debt4punto2);
                }

                // 35 RIGA
                XSSFRow row35 = sheet.createRow(34);
                XSSFCell A35 = row35.createCell(0);
                XSSFCell B35 = row35.createCell(1);
                XSSFCell C35 = row35.createCell(2);
                XSSFCell D35 = row35.createCell(3);
                XSSFCell E35 = row35.createCell(4);
                XSSFCell F35 = row35.createCell(5);
                A35.setCellValue("II)");
                A35.setCellStyle(grassetto);
                B35.setCellValue("Crediti");
                B35.setCellStyle(grassetto);
                C35.setCellFormula("C36+C39+C40+C41+C42+C43+C44+C45");
                C35.setCellStyle(grassetto);
                E35.setCellValue("5) Debiti verso altri finanziatori:");
                F35.setCellFormula("F36+F37");
                // 36 RIGA
                XSSFRow row36 = sheet.createRow(35);
                XSSFCell A36 = row36.createCell(0);
                XSSFCell B36 = row36.createCell(1);
                XSSFCell C36 = row36.createCell(2);
                XSSFCell D36 = row36.createCell(3);
                XSSFCell E36 = row36.createCell(4);
                XSSFCell F36 = row36.createCell(5);
                B36.setCellValue("1) Verso clienti");
                fatturedaemettere = findViewById(R.id.fatturedaemettere);
                crediticommdiversi = findViewById(R.id.crediticommdiversi);
                debiti5punto1=findViewById(R.id.debiti_5_1);

                crediti1punto1=findViewById(R.id.crediti_1_1);
                crediti1punto2=findViewById(R.id.crediti_1_2);

                clienticostianticipati = findViewById(R.id.clienticcostianticipati);
                cambialiattive = findViewById(R.id.crediti_cambiali_attive);
                cambialiallosconto = findViewById(R.id.cambialiallosconto);
                cambialiallincasso = findViewById(R.id.cambialiallincasso);
                cambialiinsolute = findViewById(R.id.cambialiinsolute);
                fondorischisucrediti = findViewById(R.id.fondorischisucrediti);
                fondosvalutazionecrediti = findViewById(R.id.fondosvalutazionicrediti);
                if(crediti1punto2.getText().toString().length()!=0||crediti1punto1.getText().toString().length()!=0||crediticommdiversi.getText().toString().length()!=0||fatturedaemettere.getText().toString().length()!=0||clienticostianticipati.getText().toString().length()!=0||cambialiattive.getText().toString().length()!=0||cambialiallosconto.getText().toString().length()!=0||cambialiallincasso.getText().toString().length()!=0||cambialiinsolute.getText().toString().length()!=0||fondorischisucrediti.getText().toString().length()!=0||fondosvalutazionecrediti.getText().toString().length()!=0){
                    double cr1punto2= crediti1punto2.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(crediti1punto2.getText().toString());
                    double cr1punto1= crediti1punto1.getText().toString().trim().isEmpty()?0: Double.parseDouble(crediti1punto1.getText().toString());
                    double fatturedaemetter = fatturedaemettere.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(fatturedaemettere.getText().toString());
                    double crediticomdiv = crediticommdiversi.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(crediticommdiversi.getText().toString());
                    double clienticosti = clienticostianticipati.getText().toString().trim().isEmpty()? 0: Double.parseDouble(clienticostianticipati.getText().toString());
                    double cambialiattiv = cambialiattive.getText().toString().trim().isEmpty()?0: Double.parseDouble(cambialiattive.getText().toString());
                    double cambialialloscont = cambialiallosconto.getText().toString().trim().isEmpty()?0: Double.parseDouble(cambialiallosconto.getText().toString());
                    double cambialiallincass = cambialiallincasso.getText().toString().trim().isEmpty()?0: Double.parseDouble(cambialiallincasso.getText().toString());
                    double cambialiinsolut = cambialiinsolute.getText().toString().trim().isEmpty()?0: Double.parseDouble(cambialiinsolute.getText().toString());
                    double fondorischisucredit = fondorischisucrediti.getText().toString().trim().isEmpty()?0: Double.parseDouble(fondorischisucrediti.getText().toString());
                    double fondosvalutazionecredit = fondosvalutazionecrediti.getText().toString().trim().isEmpty()?0: Double.parseDouble(fondosvalutazionecrediti.getText().toString());
                    C36.setCellValue(cr1punto2 + cr1punto1 + fatturedaemetter + crediticomdiv + clienticosti + cambialiattiv + cambialialloscont + cambialiallincass + cambialiinsolut + fondorischisucredit + fondosvalutazionecredit);
                }

                E36.setCellValue("- entro l'anno");

                if(debiti5punto1.getText().toString().length()!=0 ){
                    double debt5punto1=debiti5punto1.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(debiti5punto1.getText().toString());
                    F36.setCellValue(debt5punto1);
                }



                // 37 RIGA
                XSSFRow row37 = sheet.createRow(36);
                XSSFCell A37 = row37.createCell(0);
                XSSFCell B37 = row37.createCell(1);
                XSSFCell C37 = row37.createCell(2);
                XSSFCell D37 = row37.createCell(3);
                XSSFCell E37 = row37.createCell(4);
                XSSFCell F37 = row37.createCell(5);
                B37.setCellValue("- entro l'anno");
                E37.setCellValue("- oltre l'anno");


                if(crediti1punto1.getText().toString().length()!=0){
                    double cr1punto1= crediti1punto1.getText().toString().trim().isEmpty()?0: Double.parseDouble(crediti1punto1.getText().toString());
                    C37.setCellValue(cr1punto1);
                }

                debiti5punto2=findViewById(R.id.debiti_5_2);
                if(debiti5punto2.getText().toString().length()!=0){
                    double debt5punto2=Double.parseDouble(debiti5punto2.getText().toString());
                    F37.setCellValue(debt5punto2);
                }

                //38 RIGA
                XSSFRow row38 = sheet.createRow(37);
                XSSFCell A38 = row38.createCell(0);
                XSSFCell B38 = row38.createCell(1);
                XSSFCell C38 = row38.createCell(2);
                XSSFCell D38 = row38.createCell(3);
                XSSFCell E38 = row38.createCell(4);
                XSSFCell F38 = row38.createCell(5);
                B38.setCellValue("- oltre l'anno");
                E38.setCellValue("6) Acconti");
                
                clienticacconti = findViewById(R.id.clienti_cacconti);

                if(crediti1punto2.getText().toString().length()!=0){
                    double cr1punto2= crediti1punto2.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(crediti1punto2.getText().toString());
                    C38.setCellValue(cr1punto2);
                }

                debiti6=findViewById(R.id.debiti_6);
                if(debiti6.getText().toString().length()!=0||clienticacconti.getText().toString().length()!=0){
                    double debt6= debiti6.getText().toString().trim().isEmpty()?0: Double.parseDouble(debiti6.getText().toString());
                    double clienticacconti = clientiacconti.getText().toString().trim().isEmpty()? 0 : Double.parseDouble(clientiacconti.getText().toString());
                    F38.setCellValue(debt6+ clienticacconti);
                }

                // 39 RIGA
                XSSFRow row39 = sheet.createRow(38);
                XSSFCell A39 = row39.createCell(0);
                XSSFCell B39 = row39.createCell(1);
                XSSFCell C39 = row39.createCell(2);
                XSSFCell D39 = row39.createCell(3);
                XSSFCell E39 = row39.createCell(4);
                XSSFCell F39 = row39.createCell(5);


                crediti2 = findViewById(R.id.crediti_2);
                String crediti2Text = crediti2.getText().toString().trim();
                if ( !crediti2Text.isEmpty()) {
                    try {

                        double cr2 = crediti2Text.isEmpty() ? 0 : Double.parseDouble(crediti2Text);
                        C39.setCellValue(cr2);
                    } catch (Exception e){

                    }
                }

                fatturedaricevere = findViewById(R.id.debiti_fatture_da_ricevere);
                debiti7=findViewById(R.id.debiti_7);
                String fatturedaricevereText = fatturedaricevere.getText().toString().trim();
                String debiti7Text = debiti7.getText().toString().trim();
                if(!fatturedaricevereText.isEmpty()||!debiti7Text.isEmpty()){
                    try{
                        double debt7= debiti7Text.isEmpty()? 0 : Double.parseDouble(debiti7.getText().toString());
                        double fatturedariceveree = fatturedaricevereText.isEmpty()? 0 : Double.parseDouble(fatturedaricevereText);
                        F39.setCellValue(debt7 + fatturedariceveree);
                    } catch(Exception e){

                    }

                }
                B39.setCellValue("2) Verso imprese controllate");
                E39.setCellValue("7) Debiti verso fornitori");
                // 40 RIGA
                XSSFRow row40 = sheet.createRow(39);
                XSSFCell A40 = row40.createCell(0);
                XSSFCell B40 = row40.createCell(1);
                XSSFCell C40 = row40.createCell(2);
                XSSFCell D40 = row40.createCell(3);
                XSSFCell E40 = row40.createCell(4);
                XSSFCell F40 = row40.createCell(5);


                crediti3=findViewById(R.id.crediti_3);
                if(crediti3.getText().toString().length()!=0){
                    double cr3=Double.parseDouble(crediti3.getText().toString());
                    C40.setCellValue(cr3);
                }

                debiti8=findViewById(R.id.debiti_8);
                if(debiti8.getText().toString().length()!=0){
                    double debt8=Double.parseDouble(debiti8.getText().toString());
                    F40.setCellValue(debt8);
                }
                B40.setCellValue("3) Verso imprese collegate");
                E40.setCellValue("8) Debiti rappresentati da titoli di credito");


                XSSFRow row41 = sheet.createRow(40);
                XSSFCell A41 = row41.createCell(0);
                XSSFCell B41 = row41.createCell(1);
                XSSFCell C41 = row41.createCell(2);
                XSSFCell D41 = row41.createCell(3);
                XSSFCell E41 = row41.createCell(4);
                XSSFCell F41 = row41.createCell(5);

                crediti4=findViewById(R.id.crediti_4);
                if(crediti4.getText().toString().length()!=0){
                    double cr4=Double.parseDouble(crediti4.getText().toString());
                    C41.setCellValue(cr4);
                }

                debiti9=findViewById(R.id.debiti_9);
                if(debiti9.getText().toString().length()!=0){
                    double debt9=Double.parseDouble(debiti9.getText().toString());
                    F41.setCellValue(debt9);
                }

                B41.setCellValue("4) Verso controllanti");
                E41.setCellValue("9) Debiti verso imprese controllate");
                // 42 RIGA
                XSSFRow row42 = sheet.createRow(41);
                XSSFCell A42 = row42.createCell(0);
                XSSFCell B42 = row42.createCell(1);
                XSSFCell C42 = row42.createCell(2);
                XSSFCell D42 = row42.createCell(3);
                XSSFCell E42 = row42.createCell(4);
                XSSFCell F42 = row42.createCell(5);

               // crediti5=findViewById(R.id.crediti_5);
                if(crediti5.getText().toString().length()!=0){
                    double cr5=Double.parseDouble(crediti5.getText().toString());
                    C42.setCellValue(cr5);
                }


                debiti10=findViewById(R.id.debiti_10);
                if(debiti10.getText().toString().length()!=0){
                    double debt10=Double.parseDouble(debiti10.getText().toString());
                    F42.setCellValue(debt10);
                }
                B42.setCellValue("5) Verso imprese sottoposte al controllo delle controllanti");
                B42.setCellStyle(grassetto);
                E42.setCellValue("10) Debiti verso imprese collegate");
                // 43 RIGA
                XSSFRow row43 = sheet.createRow(42);
                XSSFCell A43 = row43.createCell(0);
                XSSFCell B43 = row43.createCell(1);
                XSSFCell C43 = row43.createCell(2);
                XSSFCell D43 = row43.createCell(3);
                XSSFCell E43 = row43.createCell(4);
                XSSFCell F43 = row43.createCell(5);


                crediti5punto2=findViewById(R.id.crediti_5_2);
                if(crediti5punto2.getText().toString().length()!=0){
                    double cr5punto2=Double.parseDouble(crediti5punto2.getText().toString());
                    C43.setCellValue(cr5punto2);
                }

                debiti11=findViewById(R.id.debiti_11);
                if(debiti11.getText().toString().length()!=0){
                    double debt11=Double.parseDouble(debiti11.getText().toString());
                    F43.setCellValue(debt11);
                }
                B43.setCellValue("5bis)crediti tributari");
                E43.setCellValue("11) Debiti verso controllanti");
                //44 RIGA
                XSSFRow row44 = sheet.createRow(43);
                XSSFCell A44 = row44.createCell(0);
                XSSFCell B44 = row44.createCell(1);
                XSSFCell C44 = row44.createCell(2);
                XSSFCell D44 = row44.createCell(3);
                XSSFCell E44 = row44.createCell(4);
                XSSFCell F44 = row44.createCell(5);


                crediti5punto3=findViewById(R.id.crediti_5_3);
                if(crediti5punto3.getText().toString().length()!=0){
                    double cr5punto3=Double.parseDouble(crediti5punto3.getText().toString());
                    C44.setCellValue(cr5punto3);
                }

                // debiti11punto2=findViewById(R.id.debiti_11_2);
                if(debiti11punto2.getText().toString().length()!=0){
                    double debt11punto2=Double.parseDouble(debiti11punto2.getText().toString());
                    F44.setCellValue(debt11punto2);
                }
                B44.setCellValue("5ter)imposte anticipate");
                E44.setCellValue("11bis) debiti verso imprese sottoposte al controllo delle controllanti");
                //45 RIGA
                XSSFRow row45 = sheet.createRow(44);
                XSSFCell A45 = row45.createCell(0);
                XSSFCell B45 = row45.createCell(1);
                XSSFCell C45 = row45.createCell(2);
                XSSFCell D45 = row45.createCell(3);
                XSSFCell E45 = row45.createCell(4);
                XSSFCell F45 = row45.createCell(5);

                crediti5punto4=findViewById(R.id.crediti_5_4);
                if(crediti5punto4.getText().toString().length()!=0){
                    double cr5punto4=Double.parseDouble(crediti5punto4.getText().toString());
                    C45.setCellValue(cr5punto4);
                }

                debiti12=findViewById(R.id.debiti_12);
                erariociv = findViewById(R.id.erario_civa);
                if(debiti12.getText().toString().length()!=0||erariociv.getText().toString().length()!=0){
                    double debt12=debiti12.getText().toString().trim().isEmpty()?0: Double.parseDouble(debiti12.getText().toString());
                    double erariociva = erariociv.getText().toString().trim().isEmpty()?0: Double.parseDouble(erariociv.getText().toString());
                    F45.setCellValue(debt12 + erariociva);
                }
                B45.setCellValue("5quater)verso altri");
                E45.setCellValue("12) Debiti tributari");
                //46 RIGA
                XSSFRow row46 = sheet.createRow(45);
                XSSFCell A46 = row46.createCell(0);
                XSSFCell B46 = row46.createCell(1);
                XSSFCell C46 = row46.createCell(2);
                XSSFCell D46 = row46.createCell(3);
                XSSFCell E46 = row46.createCell(4);
                XSSFCell F46 = row46.createCell(5);
                B46.setCellValue("III. Attività finan. non immob.");
                B46.setCellStyle(grassetto);
                E46.setCellValue("13) Debiti verso istituti di previdenza e di sicurezza sociale");
                C46.setCellFormula("SUM(C47:C54)");
                C46.setCellStyle(grassetto);


             //   debiti13=findViewById(R.id.debiti_13);
                if(debiti13.getText().toString().length()!=0){
                    double debt13=Double.parseDouble(debiti13.getText().toString());
                    F46.setCellValue(debt13);
                }
                //47 RIGA
                XSSFRow row47 = sheet.createRow(46);
                XSSFCell A47 = row47.createCell(0);
                XSSFCell B47 = row47.createCell(1);
                XSSFCell C47 = row47.createCell(2);
                XSSFCell D47 = row47.createCell(3);
                XSSFCell E47 = row47.createCell(4);
                XSSFCell F47 = row47.createCell(5);


                debiti14=findViewById(R.id.debiti_14);
                if(debiti14.getText().toString().length()!=0){
                    double debt14=Double.parseDouble(debiti14.getText().toString());
                    F47.setCellValue(debt14);
                }

                immobilizzazionifinanziarienonimmobilizzate1=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_1);
                if(immobilizzazionifinanziarienonimmobilizzate1.getText().toString().length()!=0){
                    double immfni1=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate1.getText().toString());
                    C47.setCellValue(immfni1);
                }
                B47.setCellValue("1) Partecipazioni in imprese controllate");
                E47.setCellValue("14) Altri debiti");
                //48 RIGA
                XSSFRow row48 = sheet.createRow(47);
                XSSFCell A48 = row48.createCell(0);
                XSSFCell B48 = row48.createCell(1);
                XSSFCell C48 = row48.createCell(2);
                XSSFCell D48 = row48.createCell(3);
                XSSFCell E48 = row48.createCell(4);
                XSSFCell F48 = row48.createCell(5);
                B48.setCellValue("2) Partecipazioni in imprese collegate");


                immobilizzazionifinanziarienonimmobilizzate2=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_2);
                if(immobilizzazionifinanziarienonimmobilizzate2.getText().toString().length()!=0){
                    double immfni2=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate2.getText().toString());
                    C48.setCellValue(immfni2);
                }
                //49 RIGA
                XSSFRow row49 = sheet.createRow(48);
                XSSFCell A49 = row49.createCell(0);
                XSSFCell B49 = row49.createCell(1);
                XSSFCell C49 = row49.createCell(2);
                XSSFCell D49 = row49.createCell(3);
                XSSFCell E49 = row49.createCell(4);
                XSSFCell F49 = row49.createCell(5);
                B49.setCellValue("3) Partecipazioni in imprese controllanti");


                immobilizzazionifinanziarienonimmobilizzate3=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_3);
                if(immobilizzazionifinanziarienonimmobilizzate3.getText().toString().length()!=0){
                    double immfni3=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate3.getText().toString());
                    C49.setCellValue(immfni3);
                }
                // 50 RIGA
                XSSFRow row50 = sheet.createRow(49);
                XSSFCell A50 = row50.createCell(0);
                XSSFCell B50 = row50.createCell(1);
                XSSFCell C50 = row50.createCell(2);
                XSSFCell D50 = row50.createCell(3);
                XSSFCell E50 = row50.createCell(4);
                XSSFCell F50 = row50.createCell(5);
                B50.setCellValue("3bis) partecipazioni in imprese sottoposte al controllo delle controllanti");
                B50.setCellStyle(grassetto);


              //  immobilizzazionifinanziarienonimmobilizzate3punto2=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_3_2);
                if(immobilizzazionifinanziarienonimmobilizzate3punto2.getText().toString().length()!=0){
                    double immfni3punto2=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate3punto2.getText().toString());
                    C50.setCellValue(immfni3punto2);
                }
                //51 RIGA
                XSSFRow row51 = sheet.createRow(50);
                XSSFCell A51 = row51.createCell(0);
                XSSFCell B51 = row51.createCell(1);
                XSSFCell C51 = row51.createCell(2);
                XSSFCell D51 = row51.createCell(3);
                XSSFCell E51 = row51.createCell(4);
                XSSFCell F51 = row51.createCell(5);
                B51.setCellValue("4) altre partecipazioni");


                immobilizzazionifinanziarienonimmobilizzate4=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_4);
                if(immobilizzazionifinanziarienonimmobilizzate4.getText().toString().length()!=0){
                    double immfni4=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate4.getText().toString());
                    C51.setCellValue(immfni4);
                }
                //52 RIGA
                XSSFRow row52 = sheet.createRow(51);
                XSSFCell A52 = row52.createCell(0);
                XSSFCell B52 = row52.createCell(1);
                XSSFCell C52 = row52.createCell(2);
                XSSFCell D52 = row52.createCell(3);
                XSSFCell E52 = row52.createCell(4);
                XSSFCell F52 = row52.createCell(5);
                B52.setCellValue("5) Strumenti finanziari derivati attivi");
                B52.setCellStyle(grassetto);

                immobilizzazionifinanziarienonimmobilizzate5=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_5);
                if(immobilizzazionifinanziarienonimmobilizzate5.getText().toString().length()!=0){
                    double immfni5=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate5.getText().toString());
                    C52.setCellValue(immfni5);
                }
                //53 RIGA
                XSSFRow row53 = sheet.createRow(52);
                XSSFCell A53 = row53.createCell(0);
                XSSFCell B53 = row53.createCell(1);
                XSSFCell C53 = row53.createCell(2);
                XSSFCell D53 = row53.createCell(3);
                XSSFCell E53 = row53.createCell(4);
                XSSFCell F53 = row53.createCell(5);
                B53.setCellValue("6) Altri titoli");
                XSSFRow row54 = sheet.createRow(53);

                immobilizzazionifinanziarienonimmobilizzate6=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_6);
                if(immobilizzazionifinanziarienonimmobilizzate6.getText().toString().length()!=0){
                    double immfni6=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate6.getText().toString());
                    C53.setCellValue(immfni6);
                }

                //54 RIGA
                XSSFCell A54 = row54.createCell(0);
                XSSFCell B54 = row54.createCell(1);
                XSSFCell C54 = row54.createCell(2);
                XSSFCell D54 = row54.createCell(3);
                XSSFCell E54 = row54.createCell(4);
                XSSFCell F54 = row54.createCell(5);
                B54.setCellValue("7) Altre");

                immobilizzazionifinanziarienonimmobilizzate7=findViewById(R.id.immobilizzazionifinanziarienonimmobilizzate_7);
                if(immobilizzazionifinanziarienonimmobilizzate7.getText().toString().length()!=0){
                    double immfni7=Double.parseDouble(immobilizzazionifinanziarienonimmobilizzate7.getText().toString());
                    C54.setCellValue(immfni7);
                }
                //55 RIGA
                XSSFRow row55 = sheet.createRow(54);
                XSSFCell A55 = row55.createCell(0);
                XSSFCell B55 = row55.createCell(1);
                XSSFCell C55 = row55.createCell(2);
                XSSFCell D55 = row55.createCell(3);
                XSSFCell E55 = row55.createCell(4);
                XSSFCell F55 = row55.createCell(5);
                A55.setCellValue("IV)");
                A55.setCellStyle(grassetto);
                B55.setCellValue("Disponibilità liquide");
                B55.setCellStyle(grassetto);
                C55.setCellFormula("SUM(C56:C58)");
                C55.setCellStyle(grassetto);
                //56 RIGA
                XSSFRow row56 = sheet.createRow(55);
                XSSFCell A56 = row56.createCell(0);
                XSSFCell B56 = row56.createCell(1);
                XSSFCell C56 = row56.createCell(2);
                XSSFCell D56 = row56.createCell(3);
                XSSFCell E56 = row56.createCell(4);
                XSSFCell F56 = row56.createCell(5);
                B56.setCellValue("1) Depositi bancari e postali");

                disponibilitàliquide1=findViewById(R.id.disponibilitàliquide_1);
                if(disponibilitàliquide1.getText().toString().length()!=0){
                    double displ1=Double.parseDouble(disponibilitàliquide1.getText().toString());
                    C56.setCellValue(displ1);
                }
                //57 RIGA
                XSSFRow row57 = sheet.createRow(56);
                XSSFCell A57 = row57.createCell(0);
                XSSFCell B57 = row57.createCell(1);
                XSSFCell C57 = row57.createCell(2);
                XSSFCell D57 = row57.createCell(3);
                XSSFCell E57 = row57.createCell(4);
                XSSFCell F57 = row57.createCell(5);
                B57.setCellValue("2) Assegni");
                D57.setCellValue("E)");
                D57.setCellStyle(grassetto);
                E57.setCellValue("RATEI E RISCONTI");
                E57.setCellStyle(grassetto);
                F57.setCellFormula("F58");
                F57.setCellStyle(grassetto);


                disponibilitàliquide2=findViewById(R.id.disponibilitàliquide_2);
                if(disponibilitàliquide2.getText().toString().length()!=0){
                    double displ12=Double.parseDouble(disponibilitàliquide2.getText().toString());
                    C57.setCellValue(displ12);
                }
                //58 RIGA
                XSSFRow row58 = sheet.createRow(57);
                XSSFCell A58 = row58.createCell(0);
                XSSFCell B58 = row58.createCell(1);
                XSSFCell C58 = row58.createCell(2);
                XSSFCell D58 = row58.createCell(3);
                XSSFCell E58 = row58.createCell(4);
                XSSFCell F58 = row58.createCell(5);
                B58.setCellValue("3) Denaro e valori in cassa");
                E58.setCellValue("Ratei passivi");

                rateieriscontipassivi1=findViewById(R.id.rateieriscontipassivi_1);
                if(rateieriscontipassivi1.getText().toString().length()!=0){
                    double rerp1=Double.parseDouble(rateieriscontipassivi1.getText().toString());
                    F58.setCellValue(rerp1);
                }

                disponibilitàliquide3=findViewById(R.id.disponibilitàliquide_3);
                if(disponibilitàliquide3.getText().toString().length()!=0){
                    double displ3=Double.parseDouble(disponibilitàliquide3.getText().toString());
                    C58.setCellValue(displ3);
                }
                //59 RIGA
                XSSFRow row59 = sheet.createRow(58);
                XSSFCell A59 = row59.createCell(0);
                XSSFCell B59 = row59.createCell(1);
                XSSFCell C59 = row59.createCell(2);
                XSSFCell D59 = row59.createCell(3);
                XSSFCell E59 = row59.createCell(4);
                XSSFCell F59 = row59.createCell(5);
                A59.setCellValue("D)");
                A59.setCellStyle(grassetto);
                B59.setCellValue("RATEI E RISCONTI");
                B59.setCellStyle(grassetto);
                C59.setCellStyle(grassetto);


                rateieriscontiattivi1=findViewById(R.id.rateieriscontiattivi_1);
                if(rateieriscontiattivi1.getText().toString().length()!=0){
                    double rera1=Double.parseDouble(rateieriscontiattivi1.getText().toString());
                    C59.setCellValue(rera1);
                }




                //60 RIGA
                XSSFRow row60 = sheet.createRow(59);
                XSSFCell A60 = row60.createCell(0);
                XSSFCell B60 = row60.createCell(1);
                XSSFCell C60 = row60.createCell(2);
                XSSFCell D60 = row60.createCell(3);
                XSSFCell E60 = row60.createCell(4);
                XSSFCell F60 = row60.createCell(5);
                //61 RIGA
                XSSFRow row61 = sheet.createRow(60);
                XSSFCell A61 = row61.createCell(0);
                XSSFCell B61 = row61.createCell(1);
                XSSFCell C61 = row61.createCell(2);
                XSSFCell D61 = row61.createCell(3);
                XSSFCell E61 = row61.createCell(4);
                XSSFCell F61 = row61.createCell(5);

                A61.setCellValue("TOTALE ATTIVO");
                A61.setCellStyle(alcentro);
                C61.setCellFormula("C5+C7+C28+C59");
                C61.setCellStyle(grassetto);
                D61.setCellValue("TOTALE PASSIVO");
                D61.setCellStyle(alcentro);
                E61.setCellStyle(style);
                F61.setCellFormula("F5+F18+F25+F28+F57");
                F61.setCellStyle(grassetto);


             RegionUtil.setBorderBottom(BorderStyle.MEDIUM,
                     CellRangeAddress.valueOf("A61:F61"), sheet);
             RegionUtil.setBorderTop(BorderStyle.MEDIUM,
                     CellRangeAddress.valueOf("A3:F3"), sheet);
             RegionUtil.setBorderRight(BorderStyle.MEDIUM,
                     CellRangeAddress.valueOf("C3:C61"), sheet);
             RegionUtil.setBorderRight(BorderStyle.MEDIUM,
                     CellRangeAddress.valueOf("F3:F61"), sheet);
             RegionUtil.setBorderLeft(BorderStyle.MEDIUM,
                     CellRangeAddress.valueOf("A3:A61"), sheet);
             RegionUtil.setBorderRight(sottile.getBorderRight(), CellRangeAddress.valueOf("A4:A60"), sheet);
            RegionUtil.setBorderRight(sottile.getBorderRight(), CellRangeAddress.valueOf("B4:B61"), sheet);
            RegionUtil.setBorderRight(sottile.getBorderRight(), CellRangeAddress.valueOf("D4:D60"), sheet);
            RegionUtil.setBorderRight(sottile.getBorderRight(), CellRangeAddress.valueOf("E4:E61"), sheet);
                for(int i=4; i<=61; i++) {
                    RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress.valueOf("A"+i+":F"+i), sheet);
                }


                // IMPOSTARE LARGHEZZA COLONNE
             sheet.setColumnWidth(1, 10000);
             sheet.setColumnWidth(4, 10000);

                XSSFSheet sheet2 = workbook.createSheet("SP Ricl.criterio finanziario");
                XSSFRow roww1 = sheet2.createRow(0);
                XSSFRow roww2 = sheet2.createRow(1);
                XSSFRow roww3 = sheet2.createRow(2);
                XSSFRow roww4 = sheet2.createRow(3);
                XSSFRow roww5 = sheet2.createRow(4);
                XSSFRow roww6 = sheet2.createRow(5);
                XSSFRow roww7 = sheet2.createRow(6);
                XSSFRow roww8 = sheet2.createRow(7);
                XSSFRow roww9 = sheet2.createRow(8);
                XSSFRow roww10 = sheet2.createRow(9);
                XSSFRow roww11 = sheet2.createRow(10);
                XSSFRow roww12 = sheet2.createRow(11);
                XSSFRow roww13 = sheet2.createRow(12);
                XSSFRow roww14 = sheet2.createRow(13);
                XSSFRow roww15 = sheet2.createRow(14);
                XSSFRow roww16 = sheet2.createRow(15);
                XSSFRow roww17 = sheet2.createRow(16);
                XSSFRow roww18 = sheet2.createRow(17);
                XSSFRow roww19 = sheet2.createRow(18);
                XSSFRow roww20 = sheet2.createRow(19);
                XSSFRow roww21 = sheet2.createRow(20);
                XSSFRow roww22 = sheet2.createRow(21);
                XSSFRow roww23 = sheet2.createRow(22);
                XSSFRow roww24 = sheet2.createRow(23);
                XSSFRow roww25 = sheet2.createRow(24);
                XSSFRow roww26 = sheet2.createRow(25);
                XSSFRow roww27 = sheet2.createRow(26);
                XSSFRow roww28 = sheet2.createRow(27);
                XSSFRow roww29 = sheet2.createRow(28);
                XSSFRow roww30 = sheet2.createRow(29);
                XSSFRow roww31 = sheet2.createRow(30);
                XSSFRow roww32 = sheet2.createRow(31);
                XSSFRow roww33 = sheet2.createRow(32);
                XSSFRow roww34 = sheet2.createRow(33);
                XSSFRow roww35 = sheet2.createRow(34);



                sheet2.addMergedRegion(CellRangeAddress.valueOf("A1:D1"));
                Cell aa1=roww1.createCell(0);
                Cell bb1=roww1.createCell(1);
                Cell cc1=roww1.createCell(2);
                Cell dd1=roww1.createCell(3);
                aa1.setCellValue("STATO PATRIMONIALE RICLASSIFICATO SEGUENDO IL CRITERIO FINANZIARIO");
                //2 RIGA
                Cell aa2=roww2.createCell(0);
                Cell bb2=roww2.createCell(1);
                Cell cc2=roww2.createCell(2);
                Cell dd2=roww2.createCell(3);
                //3 RIGA
                Cell aa3=roww3.createCell(0);
                Cell bb3=roww3.createCell(1);
                Cell cc3=roww3.createCell(2);
                Cell dd3=roww3.createCell(3);
                aa3.setCellValue("ATTIVO");
                bb3.setCellValue("data");
                cc3.setCellValue("PASSIVO");
                dd3.setCellValue("data");
                //4 RIGA
                Cell aa4=roww4.createCell(0);
                Cell bb4=roww4.createCell(1);
                Cell cc4=roww4.createCell(2);
                Cell dd4=roww4.createCell(3);
                bb4.setCellValue("Val.ass.");
                dd4.setCellValue("Val.ass.");
                //5 RIGA
                Cell aa5=roww5.createCell(0);
                Cell bb5=roww5.createCell(1);
                Cell cc5=roww5.createCell(2);
                Cell dd5=roww5.createCell(3);
                //6 RIGA
                Cell aa6=roww6.createCell(0);
                Cell bb6=roww6.createCell(1);
                Cell cc6=roww6.createCell(2);
                Cell dd6=roww6.createCell(3);
                aa6.setCellValue("IMMOBILIZZAZIONI");
                bb6.setCellFormula("B7+B10+B13+B16");
                cc6.setCellValue("PATRIMONIO NETTO");
                dd6.setCellFormula("SUM(D7:D17)");
                //7 RIGA
                Cell aa7=roww7.createCell(0);
                Cell bb7=roww7.createCell(1);
                Cell cc7=roww7.createCell(2);
                Cell dd7=roww7.createCell(3);
                aa7.setCellValue("Immobilizzazioni Immateriali");
                bb7.setCellFormula("SUM(B8:B9)");
                cc7.setCellValue("I. Capitale sociale");
                if(patrimonionetto1.getText().toString().length()!=0){
                    double patnetto1=Double.parseDouble(patrimonionetto1.getText().toString());
                    cellF6.setCellValue(patnetto1);
                    dd7.setCellValue(patnetto1);
                }

                //8 RIGA
                Cell aa8=roww8.createCell(0);
                Cell bb8=roww8.createCell(1);
                Cell cc8=roww8.createCell(2);
                Cell dd8=roww8.createCell(3);
                //9 RIGA
                Cell aa9=roww9.createCell(0);
                Cell bb9=roww9.createCell(1);
                Cell cc9=roww9.createCell(2);
                Cell dd9=roww9.createCell(3);
                //10 RIGA
                Cell aa10=roww10.createCell(0);
                Cell bb10=roww10.createCell(1);
                Cell cc10=roww10.createCell(2);
                Cell dd10=roww10.createCell(3);
                //11 RIGA
                Cell aa11=roww11.createCell(0);
                Cell bb11=roww11.createCell(1);
                Cell cc11=roww11.createCell(2);
                Cell dd11=roww11.createCell(3);
                //12 RIGA
                Cell aa12=roww12.createCell(0);
                Cell bb12=roww12.createCell(1);
                Cell cc12=roww12.createCell(2);
                Cell dd12=roww12.createCell(3);
                //13 RIGA
                Cell aa13=roww13.createCell(0);
                Cell bb13=roww13.createCell(1);
                Cell cc13=roww13.createCell(2);
                Cell dd13=roww13.createCell(3);
                //14 RIGA
                Cell aa14=roww14.createCell(0);
                Cell bb14=roww14.createCell(1);
                Cell cc14=roww14.createCell(2);
                Cell dd14=roww14.createCell(3);
                //15 RIGA
                Cell aa15=roww15.createCell(0);
                Cell bb15=roww15.createCell(1);
                Cell cc15=roww15.createCell(2);
                Cell dd15=roww15.createCell(3);
                //16 RIGA
                Cell aa16=roww16.createCell(0);
                Cell bb16=roww16.createCell(1);
                Cell cc16=roww16.createCell(2);
                Cell dd16=roww16.createCell(3);
                //17 RIGA
                Cell aa17=roww17.createCell(0);
                Cell bb17=roww17.createCell(1);
                Cell cc17=roww17.createCell(2);
                Cell dd17=roww17.createCell(3);
                //18 RIGA
                Cell aa18=roww18.createCell(0);
                Cell bb18=roww18.createCell(1);
                Cell cc18=roww18.createCell(2);
                Cell dd18=roww18.createCell(3);
                //19 RIGA
                Cell aa19=roww19.createCell(0);
                Cell bb19=roww19.createCell(1);
                Cell cc19=roww19.createCell(2);
                Cell dd19=roww19.createCell(3);
                //20 RIGA
                Cell aa20=roww20.createCell(0);
                Cell bb20=roww20.createCell(1);
                Cell cc20=roww20.createCell(2);
                Cell dd20=roww20.createCell(3);
                //21 RIGA
                Cell aa21=roww21.createCell(0);
                Cell bb21=roww21.createCell(1);
                Cell cc21=roww21.createCell(2);
                Cell dd21=roww21.createCell(3);
                //22 RIGA
                Cell aa22=roww22.createCell(0);
                Cell bb22=roww22.createCell(1);
                Cell cc22=roww22.createCell(2);
                Cell dd22=roww22.createCell(3);
                //23 RIGA
                Cell aa23=roww23.createCell(0);
                Cell bb23=roww23.createCell(1);
                Cell cc23=roww23.createCell(2);
                Cell dd23=roww23.createCell(3);
                //24 RIGA
                Cell aa24=roww24.createCell(0);
                Cell bb24=roww24.createCell(1);
                Cell cc24=roww24.createCell(2);
                Cell dd24=roww24.createCell(3);
                //25 RIGA
                Cell aa25=roww25.createCell(0);
                Cell bb25=roww25.createCell(1);
                Cell cc25=roww25.createCell(2);
                Cell dd25=roww25.createCell(3);
                //26 RIGA
                Cell aa26=roww26.createCell(0);
                Cell bb26=roww26.createCell(1);
                Cell cc26=roww26.createCell(2);
                Cell dd26=roww26.createCell(3);
                //27 RIGA
                Cell aa27=roww27.createCell(0);
                Cell bb27=roww27.createCell(1);
                Cell cc27=roww27.createCell(2);
                Cell dd27=roww27.createCell(3);
                //28 RIGA
                Cell aa28=roww28.createCell(0);
                Cell bb28=roww28.createCell(1);
                Cell cc28=roww28.createCell(2);
                Cell dd28=roww28.createCell(3);
                //29 RIGA
                Cell aa29=roww29.createCell(0);
                Cell bb29=roww29.createCell(1);
                Cell cc29=roww29.createCell(2);
                Cell dd29=roww29.createCell(3);
                //30 RIGA
                Cell aa30=roww30.createCell(0);
                Cell bb30=roww30.createCell(1);
                Cell cc30=roww30.createCell(2);
                Cell dd30=roww30.createCell(3);
                //31 RIGA
                Cell aa31=roww31.createCell(0);
                Cell bb31=roww31.createCell(1);
                Cell cc31=roww31.createCell(2);
                Cell dd31=roww31.createCell(3);
                //32 RIGA
                Cell aa32=roww32.createCell(0);
                Cell bb32=roww32.createCell(1);
                Cell cc32=roww32.createCell(2);
                Cell dd32=roww32.createCell(3);
                //33 RIGA
                Cell aa33=roww33.createCell(0);
                Cell bb33=roww33.createCell(1);
                Cell cc33=roww33.createCell(2);
                Cell dd33=roww33.createCell(3);
                //34 RIGA
                Cell aa34=roww34.createCell(0);
                Cell bb34=roww34.createCell(1);
                Cell cc34=roww34.createCell(2);
                Cell dd34=roww34.createCell(3);
                //35 RIGA
                Cell aa35=roww35.createCell(0);
                Cell bb35=roww35.createCell(1);
                Cell cc35=roww35.createCell(2);
                Cell dd35=roww35.createCell(3);



            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
             evaluator.evaluateAll();
             fileSaver.execute(workbook);
            }

        });
    }
}
