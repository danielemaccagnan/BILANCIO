package com.example.bilancio;

import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.view.MenuItem;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.PopupMenu;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import org.apache.poi.ss.usermodel.BorderStyle;
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

public class contoeconomico extends AppCompatActivity {
    public static double ricavi=0;

    Context context = this;
    private static final int REQUEST_CODE_WRITE_EXTERNAL_STORAGE = 1;

    // Dichiarazione del metodo showPopupMenu()
    private void showPopupMenu(View v) {
        PopupMenu popup = new PopupMenu(this, v);
        popup.show();
    }

    public EditText r5;
    public EditText r6;
    public EditText r7;
    public EditText r8;
    public EditText r9;
    public EditText r10;
    public EditText r11;
    public EditText r12;
    public EditText r13;
    public EditText r14;
    public EditText r16;
    public EditText r17;
    public EditText r18;
    public EditText r20;
    public EditText r21;
    public EditText r22;
    public EditText r23;
    public EditText r24;
    public EditText r25;
    public EditText r26;
    public EditText r27;
    public EditText r32;
    public EditText r33;
    public EditText r34;
    public EditText r35;
    public EditText r36;
    public EditText r37;
    public EditText r38;
    public EditText r41;
    public EditText r42;
    public EditText r43;
    public EditText r44;
    public EditText r46;
    public EditText r47;
    public EditText r48;
    public EditText r49;
    public EditText r55;
    public EditText r56;
    public EditText r59;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        // inizializza l'istanza della classe ExcelFileSaver
        ContoEconomicoSaver fileSaver = new ContoEconomicoSaver(this);
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_contoeconomico);


        // Trova il pulsante con id "popup_button" nella tua attività
        Button popup = findViewById(R.id.popup_button_ce);

// Aggiungi un listener di click al pulsante
        popup.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                // Crea un nuovo oggetto PopupMenu e lo mostra
                PopupMenu popupMenu = new PopupMenu(contoeconomico.this, view);
                popupMenu.getMenuInflater().inflate(R.menu.popup_menu_ce, popupMenu.getMenu());

                // Aggiungi un listener per gestire il click degli elementi del menu
                popupMenu.setOnMenuItemClickListener(new PopupMenu.OnMenuItemClickListener() {
                    @Override
                    public boolean onMenuItemClick(MenuItem item) {
                        switch (item.getItemId()) {
                            case R.id.menu_item_conto_economico:
                                // Avvia la nuova activity ContoEconomicoActivity quando l'utente seleziona l'opzione Conto Economico
                                Intent intent = new Intent(contoeconomico.this, statopatrimoniale.class);
                                startActivity(intent);
                                return true;
                            case R.id.menu_item_home:
                                // Torna alla home
                                Intent intent1 = new Intent(contoeconomico.this, MainActivity.class);
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

        Button salvace = findViewById(R.id.salvace);

        salvace.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                // crea una nuova istanza di ExcelFileSaver e esegui il salvataggio
                ContoEconomicoSaver fileSaver = new ContoEconomicoSaver(context);

                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Conto Economico");

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


                sheet.addMergedRegion(CellRangeAddress.valueOf("A1:C1"));
                XSSFRow row1 = sheet.createRow(0);
                XSSFCell A1 = row1.createCell(0);
                XSSFCell B1 = row1.createCell(1);
                XSSFCell C1 = row1.createCell(2);
                A1.setCellValue("CONTO ECONOMICO");
                A1.setCellStyle(alcentro);
                //2 RIGA
                XSSFRow row2 = sheet.createRow(1);
                XSSFCell A2 = row2.createCell(0);
                XSSFCell B2 = row2.createCell(1);
                XSSFCell C2 = row2.createCell(2);

                XSSFRow row3 = sheet.createRow(2);
                XSSFCell A3 = row3.createCell(0);
                XSSFCell B3 = row3.createCell(1);
                XSSFCell C3 = row3.createCell(2);
                A3.setCellValue("A)");
                A3.setCellStyle(grassetto);
                B3.setCellValue("VALORE DELLA PRODUZIONE");
                B3.setCellStyle(grassetto);
                C3.setCellFormula("C4+C5+C6+C7+C8");
                C3.setCellStyle(grassetto);

                XSSFRow row4 = sheet.createRow(3);
                XSSFCell A4 = row4.createCell(0);
                XSSFCell B4 = row4.createCell(1);
                XSSFCell C4 = row4.createCell(2);
                A4.setCellValue("1)");
                B4.setCellValue("Ricavi delle vendite e delle prestazioni");

                r5 = findViewById(R.id.edittext_ricavi);
                if(r5.getText().toString().length()!=0){
                     ricavi = Double.parseDouble(r5.getText().toString());
                    C4.setCellValue(ricavi);
                }


                XSSFRow row5 = sheet.createRow(4);
                XSSFCell A5 = row5.createCell(0);
                XSSFCell B5 = row5.createCell(1);
                XSSFCell C5 = row5.createCell(2);
                A5.setCellValue("2");
                B5.setCellValue("Variazione rimanenze di prodotti in corso di lavorazione, semil. e finiti");
                r6=findViewById(R.id.edittext_variazione_rimanenze);
                if(r6.getText().toString().length()!=0){
                    double riga6=Double.parseDouble(r6.getText().toString());
                    C5.setCellValue(riga6);
                }
                XSSFRow row6 = sheet.createRow(5);
                XSSFCell A6 = row6.createCell(0);
                XSSFCell B6 = row6.createCell(1);
                XSSFCell C6 = row6.createCell(2);
                A6.setCellValue("3)");
                B6.setCellValue("Variazione dei lavori in corso su ordinazione");
                r7=findViewById(R.id.edittext_variazione_lavori_in_corso);
                if(r7.getText().toString().length()!=0){
                    double riga7=Double.parseDouble(r7.getText().toString());
                    C6.setCellValue(riga7);
                }
                XSSFRow row7 = sheet.createRow(6);
                XSSFCell A7 = row7.createCell(0);
                XSSFCell B7 = row7.createCell(1);
                XSSFCell C7 = row7.createCell(2);
                A7.setCellValue("4)");
                B7.setCellValue("Incrementi di immobilizzazioni per lavori interni");
                r8=findViewById(R.id.edittext_incrementi_immobilizzazioni);
                if(r8.getText().toString().length()!=0){
                    double riga8=Double.parseDouble(r8.getText().toString());
                    C7.setCellValue(riga8);
                }
                XSSFRow row8 = sheet.createRow(7);
                XSSFCell A8 = row8.createCell(0);
                XSSFCell B8 = row8.createCell(1);
                XSSFCell C8 = row8.createCell(2);
                A8.setCellValue("5)");
                B8.setCellValue("Altri ricavi e proventi");
                r9=findViewById(R.id.edittext_altri_ricavi);
                if(r9.getText().toString().length()!=0){
                    double riga9=Double.parseDouble(r9.getText().toString());
                    C8.setCellValue(riga9);
                }

                XSSFRow row9 = sheet.createRow(8);
                XSSFCell A9 = row9.createCell(0);
                XSSFCell B9 = row9.createCell(1);
                XSSFCell C9 = row9.createCell(2);
                A9.setCellValue("B)");
                A9.setCellStyle(grassetto);
                B9.setCellValue("COSTI DELLA PRODUZIONE");
                B9.setCellStyle(grassetto);
                C9.setCellFormula("C10+C11+C12+C13+C19+C25+C24+C26+C27");
                C9.setCellStyle(grassetto);
//10 RIGA
                XSSFRow row10 = sheet.createRow(9);
                XSSFCell A10 = row10.createCell(0);
                XSSFCell B10 = row10.createCell(1);
                XSSFCell C10 = row10.createCell(2);
                A10.setCellValue("6)");
                B10.setCellValue("Costi per acquisti di materie prime, sussidiarie, di consumo e di merci");
                r10=findViewById(R.id.edittext_costi_per_acquisti);
                if(r10.getText().toString().length()!=0){
                    double riga10=Double.parseDouble(r10.getText().toString());
                    C10.setCellValue(riga10);
                }

                //11 RIGA
                XSSFRow row11 = sheet.createRow(10);
                XSSFCell A11 = row11.createCell(0);
                XSSFCell B11 = row11.createCell(1);
                XSSFCell C11 = row11.createCell(2);
                A11.setCellValue("7)");
                B11.setCellValue("Costi per servizi");
                r11=findViewById(R.id.edittext_costi_per_servizi);
                if(r11.getText().toString().length()!=0){
                    double riga11=Double.parseDouble(r11.getText().toString());
                    C11.setCellValue(riga11);
                }

//12 RIGA
                XSSFRow row12 = sheet.createRow(11);
                XSSFCell A12 = row12.createCell(0);
                XSSFCell B12 = row12.createCell(1);
                XSSFCell C12 = row12.createCell(2);
                A12.setCellValue("8)");
                B12.setCellValue("Costi per godimento di beni di terzi");
                r12=findViewById(R.id.edittext_costi_per_godimento);
                if(r12.getText().toString().length()!=0){
                    double riga12=Double.parseDouble(r12.getText().toString());
                    C12.setCellValue(riga12);
                }

//13 RIGA
                XSSFRow row13 = sheet.createRow(12);
                XSSFCell A13 = row13.createCell(0);
                XSSFCell B13 = row13.createCell(1);
                XSSFCell C13 = row13.createCell(2);
                A13.setCellValue("9");
                B13.setCellValue("Costi per il personale");
                B13.setCellStyle(grassetto);
                C13.setCellFormula("C14+C15+C16+C17+C18");
                C13.setCellStyle(grassetto);

//14 RIGA
                XSSFRow row14 = sheet.createRow(13);
                XSSFCell A14 = row14.createCell(0);
                XSSFCell B14 = row14.createCell(1);
                XSSFCell C14 = row14.createCell(2);
                B14.setCellValue("a) Salari e stipendi");
                r13=findViewById(R.id.edittext_salari);
                if(r13.getText().toString().length()!=0){
                    double salari = Double.parseDouble(r13.getText().toString());
                    C14.setCellValue(salari);
                }

//15 RIGA
                XSSFRow row15 = sheet.createRow(14);
                XSSFCell A15 = row15.createCell(0);
                XSSFCell B15 = row15.createCell(1);
                XSSFCell C15 = row15.createCell(2);
                B15.setCellValue("b) Oneri sociali");
                r14=findViewById(R.id.edittext_oneri_sociali);
                if(r14.getText().toString().length()!=0){
                    double oneri = Double.parseDouble(r14.getText().toString());
                    C15.setCellValue(oneri);
                }

//16 RIGA
                XSSFRow row16 = sheet.createRow(15);
                XSSFCell A16 = row16.createCell(0);
                XSSFCell B16 = row16.createCell(1);
                XSSFCell C16 = row16.createCell(2);
                B16.setCellValue("c) Trattamento di fine rapporto");
                r16=findViewById(R.id.edittext_trattamento_fine_rapporto);
                if(r16.getText().toString().length()!=0){
                    double tfr = Double.parseDouble(r16.getText().toString());
                    C16.setCellValue(tfr);
                }

//17 RIGA
                XSSFRow row17 = sheet.createRow(16);
                XSSFCell A17 = row17.createCell(0);
                XSSFCell B17 = row17.createCell(1);
                XSSFCell C17 = row17.createCell(2);
                B17.setCellValue("d) Trattamento di quiescenza e simili");
                r17=findViewById(R.id.edittext_trattamento_quiescenza);
                if(r17.getText().toString().length()!=0){
                    double quiescenza = Double.parseDouble(r17.getText().toString());
                    C17.setCellValue(quiescenza);
                }
                //18 RIGA

                XSSFRow row18 = sheet.createRow(17);
                XSSFCell A18 = row18.createCell(0);
                XSSFCell B18 = row18.createCell(1);
                XSSFCell C18 = row18.createCell(2);
                B18.setCellValue("e) Altri costi");
                r18=findViewById(R.id.edittext_altri_costi);
                if(r18.getText().toString().length()!=0){
                    double altricosti = Double.parseDouble(r18.getText().toString());
                    C18.setCellValue(altricosti);
                }

//19 RIGA
                XSSFRow row19 = sheet.createRow(18);
                XSSFCell A19 = row19.createCell(0);
                XSSFCell B19 = row19.createCell(1);
                XSSFCell C19 = row19.createCell(2);
                A19.setCellValue("10)");
                B19.setCellValue("Ammortamenti e svalutazioni");
                B19.setCellStyle(grassetto);
                C19.setCellFormula("C20+C21+C22+C23");
                C19.setCellStyle(grassetto);

//20 RIGA
                XSSFRow row20 = sheet.createRow(19);
                XSSFCell A20 = row20.createCell(0);
                XSSFCell B20 = row20.createCell(1);
                XSSFCell C20 = row20.createCell(2);
                B20.setCellValue("a) Ammortamento immobilizzazioni immateriali");
                r20=findViewById(R.id.edittext_ammortamento_immobilizzazioni_immateriali);
                if(r20.getText().toString().length()!=0){
                    double ammortamento = Double.parseDouble(r20.getText().toString());
                    C20.setCellValue(ammortamento);
                }
//21 RIGA
                XSSFRow row21 = sheet.createRow(20);
                XSSFCell A21 = row21.createCell(0);
                XSSFCell B21 = row21.createCell(1);
                XSSFCell C21 = row21.createCell(2);
                B21.setCellValue("b) Ammortamento immobilizzazioni materiali");
                r21=findViewById(R.id.edittext_ammortamento_immobilizzazioni_materiali);
                if(r21.getText().toString().length()!=0){
                    double ammortamentob = Double.parseDouble(r21.getText().toString());
                    C21.setCellValue(ammortamentob);
                }

//22 RIGA
                XSSFRow row22 = sheet.createRow(21);
                XSSFCell A22 = row22.createCell(0);
                XSSFCell B22 = row22.createCell(1);
                XSSFCell C22 = row22.createCell(2);
                B22.setCellValue("c) Altre svalutazioni delle immobilizzazioni");
                r22=findViewById(R.id.edittext_altre_svalutazioni_immobilizzazioni);
                if(r22.getText().toString().length()!=0){
                    double ammortamentoc = Double.parseDouble(r22.getText().toString());
                    C22.setCellValue(ammortamentoc);
                }


                //23 RIGA
                XSSFRow row23 = sheet.createRow(22);
                XSSFCell A23 = row23.createCell(0);
                XSSFCell B23 = row23.createCell(1);
                XSSFCell C23 = row23.createCell(2);
                B23.setCellValue("d) Svalutazione dei crediti compresi all'attivo circolante e delle disp. Liq.");
                r23=findViewById(R.id.edittext_svalutazione_crediti);
                if(r23.getText().toString().length()!=0){
                    double ammortamentod = Double.parseDouble(r23.getText().toString());
                    C23.setCellValue(ammortamentod);
                }

//24 RIGA
                XSSFRow row24 = sheet.createRow(23);
                XSSFCell A24 = row24.createCell(0);
                XSSFCell B24 = row24.createCell(1);
                XSSFCell C24 = row24.createCell(2);
                A24.setCellValue("11)");
                B24.setCellValue("Variazione delle rimanenze di materie prime, sussidiarie, di consumo e di merci");
                r24=findViewById(R.id.edittext_variazione_rimanenzee);
                if(r24.getText().toString().length()!=0){
                    double rimanenzeee = Double.parseDouble(r24.getText().toString());
                    C24.setCellValue(rimanenzeee);
                }
//25 RIGA
                XSSFRow row25 = sheet.createRow(24);
                XSSFCell A25 = row25.createCell(0);
                XSSFCell B25 = row25.createCell(1);
                XSSFCell C25 = row25.createCell(2);
                A25.setCellValue("12)");
                B25.setCellValue("Accantonamenti per rischi");
                r25=findViewById(R.id.edittext_accantonamenti_rischi);
                if(r25.getText().toString().length()!=0){
                    double rischi = Double.parseDouble(r25.getText().toString());
                    C25.setCellValue(rischi);
                }
//26 RIGA
                XSSFRow row26 = sheet.createRow(25);
                XSSFCell A26 = row26.createCell(0);
                XSSFCell B26 = row26.createCell(1);
                XSSFCell C26 = row26.createCell(2);
                A26.setCellValue("13)");
                B26.setCellValue("Altri accantonamenti");
                r26=findViewById(R.id.edittext_altri_accantonamenti);
                if(r26.getText().toString().length()!=0){
                    double altririschi = Double.parseDouble(r26.getText().toString());
                    C26.setCellValue(altririschi);
                }
//27 RIGA
                XSSFRow row27 = sheet.createRow(26);
                XSSFCell A27 = row27.createCell(0);
                XSSFCell B27 = row27.createCell(1);
                XSSFCell C27 = row27.createCell(2);
                A27.setCellValue("14)");
                B27.setCellValue("Oneri diversi di gestione");
                r27=findViewById(R.id.edittext_oneri_diversi_gestione);
                if(r27.getText().toString().length()!=0){
                    double oneridiversi = Double.parseDouble(r27.getText().toString());
                    C27.setCellValue(oneridiversi);
                }


//28 RIGA
                XSSFRow row28 = sheet.createRow(27);
                XSSFCell A28 = row28.createCell(0);
                XSSFCell B28 = row28.createCell(1);
                XSSFCell C28 = row28.createCell(2);
                B28.setCellValue("Differenza tra valore e costi della produzione (A-B)");
                B28.setCellStyle(grassetto);
                C28.setCellFormula("C3-C9");
                C28.setCellStyle(grassetto);


//29 RIGA
                XSSFRow row29 = sheet.createRow(28);
                XSSFCell A29 = row29.createCell(0);
                XSSFCell B29 = row29.createCell(1);
                XSSFCell C29 = row29.createCell(2);

//30 RIGA
                XSSFRow row30 = sheet.createRow(29);
                XSSFCell A30 = row30.createCell(0);
                XSSFCell B30 = row30.createCell(1);
                XSSFCell C30 = row30.createCell(2);
                A30.setCellValue("C)");
                A30.setCellStyle(grassetto);
                B30.setCellValue("PROVENTI E ONERI FINANZIARI");
                B30.setCellStyle(grassetto);
                C30.setCellFormula("C31+C37+C38");
                C30.setCellStyle(grassetto);

//31 RIGA
                XSSFRow row31 = sheet.createRow(30);
                XSSFCell A31 = row31.createCell(0);
                XSSFCell B31 = row31.createCell(1);
                XSSFCell C31 = row31.createCell(2);
                A31.setCellValue("15)");
                B31.setCellValue("Proventi delle partecipazioni");
                B31.setCellStyle(grassetto);
                C31.setCellFormula("C32+C33+C34+C35+C36");
                C31.setCellStyle(grassetto);
//32 RIGA
                XSSFRow row32 = sheet.createRow(31);
                XSSFCell A32 = row32.createCell(0);
                XSSFCell B32 = row32.createCell(1);
                XSSFCell C32 = row32.createCell(2);
                B32.setCellValue("a) in imprese controllate");
                r32=findViewById(R.id.edittext_imp_controllate);
                if(r32.getText().toString().length()!=0){
                    double riga32 = Double.parseDouble(r32.getText().toString());
                    C32.setCellValue(riga32);
                }


//33 RIGA
                XSSFRow row33 = sheet.createRow(32);
                XSSFCell A33 = row33.createCell(0);
                XSSFCell B33 = row33.createCell(1);
                XSSFCell C33 = row33.createCell(2);
                B33.setCellValue("b) in imprese collegate");
                r33=findViewById(R.id.edittext_imp_collegate);
                if(r33.getText().toString().length()!=0){
                    double riga33 = Double.parseDouble(r33.getText().toString());
                    C33.setCellValue(riga33);
                }


                // Creating row 34
                XSSFRow row34 = sheet.createRow(33);
                XSSFCell A34 = row34.createCell(0);
                XSSFCell B34 = row34.createCell(1);
                XSSFCell C34 = row34.createCell(2);
                B34.setCellValue("c) in imprese controllanti");
                r34=findViewById(R.id.edittext_imp_controllanti);
                if(r34.getText().toString().length()!=0){
                    double riga34 = Double.parseDouble(r34.getText().toString());
                    C34.setCellValue(riga34);
                }


// Creating row 35
                XSSFRow row35 = sheet.createRow(34);
                XSSFCell A35 = row35.createCell(0);
                XSSFCell B35 = row35.createCell(1);
                XSSFCell C35 = row35.createCell(2);
                B35.setCellValue("d) in imprese sottoposte al controllo delle controllanti");
                r35=findViewById(R.id.edittext_imp_sottoposte_controllo);
                 if(r35.getText().toString().length()!=0){
                    double riga35 = Double.parseDouble(r35.getText().toString());
                    C35.setCellValue(riga35);
                }

// Creating row 36
                XSSFRow row36 = sheet.createRow(35);
                XSSFCell A36 = row36.createCell(0);
                XSSFCell B36 = row36.createCell(1);
                XSSFCell C36 = row36.createCell(2);
                B36.setCellValue("e) in altre imprese");
                r36=findViewById(R.id.edittext_altre_imprese);
                if(r36.getText().toString().length()!=0){
                    double riga36 = Double.parseDouble(r36.getText().toString());
                    C36.setCellValue(riga36);
                }

// Creating row 37
                XSSFRow row37 = sheet.createRow(36);
                XSSFCell A37 = row37.createCell(0);
                XSSFCell B37 = row37.createCell(1);
                XSSFCell C37 = row37.createCell(2);
                A37.setCellValue("16)");
                B37.setCellValue("Altri proventi finanziari");
                r37=findViewById(R.id.edittext_altri_proventi_finanziari);
                if(r37.getText().toString().length()!=0){
                    double riga37 = Double.parseDouble(r37.getText().toString());
                    C37.setCellValue(riga37);
                }

// Creating row 38
                XSSFRow row38 = sheet.createRow(37);
                XSSFCell A38 = row38.createCell(0);
                XSSFCell B38 = row38.createCell(1);
                XSSFCell C38 = row38.createCell(2);
                A38.setCellValue("17)");
                B38.setCellValue("Interessi ed altri oneri finanziari");
                r38=findViewById(R.id.edittext_interessi_oneri_finanziari);
                if(r38.getText().toString().length()!=0){
                    double riga38 = Double.parseDouble(r38.getText().toString());
                    C38.setCellValue(riga38);
                }

// Creating row 39
                XSSFRow row39 = sheet.createRow(38);
                XSSFCell A39 = row39.createCell(0);
                XSSFCell B39 = row39.createCell(1);
                XSSFCell C39 = row39.createCell(2);
                A39.setCellValue("D)");
                A39.setCellStyle(grassetto);
                B39.setCellValue("RETTIFICHE DI VALORE DI ATTIVITA' FINANZIARIE");
                B39.setCellStyle(grassetto);

// Creating row 40
                XSSFRow row40 = sheet.createRow(39);
                XSSFCell A40 = row40.createCell(0);
                XSSFCell B40 = row40.createCell(1);
                XSSFCell C40 = row40.createCell(2);
                A40.setCellValue("18)");
                B40.setCellValue("rivalutazione di attività finanziarie");
                C40.setCellFormula("C41+42+43+44");
                B40.setCellStyle(grassetto);
                C40.setCellStyle(grassetto);

// Creating row 41
                XSSFRow row41 = sheet.createRow(40);
                XSSFCell A41 = row41.createCell(0);
                XSSFCell B41 = row41.createCell(1);
                XSSFCell C41 = row41.createCell(2);
                B41.setCellValue("a) di partecipazioni ");
                r41=findViewById(R.id.edittext_partecipazioni);
                if(r41.getText().toString().length()!=0){
                    double riga41 = Double.parseDouble(r41.getText().toString());
                    C41.setCellValue(riga41);
                }

// Creating row 42
                XSSFRow row42 = sheet.createRow(41);
                XSSFCell A42 = row42.createCell(0);
                XSSFCell B42 = row42.createCell(1);
                XSSFCell C42 = row42.createCell(2);
                B42.setCellValue("b) di immobilizzazioni finanziarie che non sono partecipazioni");
                r42=findViewById(R.id.edittext_immobili_finanziari);
                if(r42.getText().toString().length()!=0){
                    double riga42 = Double.parseDouble(r42.getText().toString());
                    C42.setCellValue(riga42);
                }


// Creating row 43
                XSSFRow row43 = sheet.createRow(42);
                XSSFCell A43 = row43.createCell(0);
                XSSFCell B43 = row43.createCell(1);
                XSSFCell C43 = row43.createCell(2);
                B43.setCellValue("c) di titoli iscritti nell'attivo circolante che non sono partecipazioni");
                r43=findViewById(R.id.edittext_titoli_attivo_circolante);
                if(r43.getText().toString().length()!=0){
                    double riga43 = Double.parseDouble(r43.getText().toString());
                    C43.setCellValue(riga43);
                }

// Creating row 44
                XSSFRow row44 = sheet.createRow(43);
                XSSFCell A44 = row44.createCell(0);
                XSSFCell B44 = row44.createCell(1);
                XSSFCell C44 = row44.createCell(2);
                B44.setCellValue("d) di strumenti finanziari derivati");
                r44=findViewById(R.id.edittext_strumenti_finanziari_derivati);
                if(r44.getText().toString().length()!=0){
                    double riga44 = Double.parseDouble(r44.getText().toString());
                    C44.setCellValue(riga44);
                }

// Creating row 45
                XSSFRow row45 = sheet.createRow(44);
                XSSFCell A45 = row45.createCell(0);
                XSSFCell B45 = row45.createCell(1);
                XSSFCell C45 = row45.createCell(2);
                A45.setCellValue("19)");
                B45.setCellValue("svalutazioni di attività finanziarie");
                B45.setCellStyle(grassetto);
                C45.setCellFormula("C46+47+48+49");
                C45.setCellStyle(grassetto);


// Creating row 46
                XSSFRow row46 = sheet.createRow(45);
                XSSFCell A46 = row46.createCell(0);
                XSSFCell B46 = row46.createCell(1);
                XSSFCell C46 = row46.createCell(2);
                B46.setCellValue("a) di partecipazioni ");
                r46=findViewById(R.id.edittext_partecipazioni_2);
                if(r46.getText().toString().length()!=0){
                    double riga46 = Double.parseDouble(r46.getText().toString());
                    C46.setCellValue(riga46);
                }


// Creating row 47
                XSSFRow row47 = sheet.createRow(46);
                XSSFCell A47 = row47.createCell(0);
                XSSFCell B47 = row47.createCell(1);
                XSSFCell C47 = row47.createCell(2);
                B47.setCellValue("b) di immobilizzazioni finanziarie che non sono partecipazioni");
                r47=findViewById(R.id.edittext_immobili_finanziari_2);
                if(r47.getText().toString().length()!=0){
                    double riga47 = Double.parseDouble(r47.getText().toString());
                    C47.setCellValue(riga47);
                }

// Creating row 48
                XSSFRow row48 = sheet.createRow(47);
                XSSFCell A48 = row48.createCell(0);
                XSSFCell B48 = row48.createCell(1);
                XSSFCell C48 = row48.createCell(2);
                B48.setCellValue("c) di titoli iscritti nell'attivo circolante che non sono partecipazioni");
                r48=findViewById(R.id.edittext_titoli_attivo_circolante_2);
                if(r48.getText().toString().length()!=0){
                    double riga48 = Double.parseDouble(r48.getText().toString());
                    C48.setCellValue(riga48);
                }

// Creating row 49
                XSSFRow row49 = sheet.createRow(48);
                XSSFCell A49 = row49.createCell(0);
                XSSFCell B49 = row49.createCell(1);
                XSSFCell C49 = row49.createCell(2);
                B49.setCellValue("d) di strumenti finanziari derivati");
                r49=findViewById(R.id.edittext_strumenti_finanziari_derivati_2);
                if(r49.getText().toString().length()!=0){
                    double riga49 = Double.parseDouble(r49.getText().toString());
                    C49.setCellValue(riga49);
                }

// Creating row 50
                XSSFRow row50 = sheet.createRow(49);
                XSSFCell A50 = row50.createCell(0);
                XSSFCell B50 = row50.createCell(1);
                XSSFCell C50 = row50.createCell(2);

// Creating row 51
                XSSFRow row51 = sheet.createRow(50);
                XSSFCell A51 = row51.createCell(0);
                XSSFCell B51 = row51.createCell(1);
                XSSFCell C51 = row51.createCell(2);
                B51.setCellValue("Totale delle rettifiche (18-19)");
                B51.setCellStyle(grassetto);
                C51.setCellFormula("C40-C45");
                C51.setCellStyle(grassetto);


// Creating row 52
                XSSFRow row52 = sheet.createRow(51);
                XSSFCell A52 = row52.createCell(0);
                XSSFCell B52 = row52.createCell(1);
                XSSFCell C52 = row52.createCell(2);

// Creating row 53
                XSSFRow row53 = sheet.createRow(52);
                XSSFCell A53 = row53.createCell(0);
                XSSFCell B53 = row53.createCell(1);
                XSSFCell C53 = row53.createCell(2);
                A53.setCellValue("E)");
                A53.setCellStyle(grassetto);
                B53.setCellValue("PROVENTI E ONERI STRAORDINARI");
                B53.setCellStyle(grassetto);
                C53.setCellFormula("C55+C56");
                C53.setCellStyle(grassetto);

// Creating row 54
                XSSFRow row54 = sheet.createRow(53);
                XSSFCell A54 = row54.createCell(0);
                XSSFCell B54 = row54.createCell(1);
                XSSFCell C54 = row54.createCell(2);

// Creating row 55
                XSSFRow row55 = sheet.createRow(54);
                XSSFCell A55 = row55.createCell(0);
                XSSFCell B55 = row55.createCell(1);
                XSSFCell C55 = row55.createCell(2);
                A55.setCellValue("20)");
                B55.setCellValue("Proventi straordinari");
                r55=findViewById(R.id.proventi_straordinari);
                if(r55.getText().toString().length()!=0){
                    double riga55 = Double.parseDouble(r55.getText().toString());
                    C55.setCellValue(riga55);
                }

// Creating row 56
                XSSFRow row56 = sheet.createRow(55);
                XSSFCell A56 = row56.createCell(0);
                XSSFCell B56 = row56.createCell(1);
                XSSFCell C56 = row56.createCell(2);
                A56.setCellValue("21)");
                B56.setCellValue("Oneri straordinari");
                r56=findViewById(R.id.oneri_straordinari);
                if(r56.getText().toString().length()!=0){
                    double riga56= Double.parseDouble(r56.getText().toString());
                    C56.setCellValue(riga56);
                }

// Creating row 57
                XSSFRow row57 = sheet.createRow(56);
                XSSFCell A57 = row57.createCell(0);
                XSSFCell B57 = row57.createCell(1);
                XSSFCell C57 = row57.createCell(2);

// Creating row 58
                XSSFRow row58 = sheet.createRow(57);
                XSSFCell A58 = row58.createCell(0);
                XSSFCell B58 = row58.createCell(1);
                XSSFCell C58 = row58.createCell(2);
                B58.setCellValue("Risultato prima delle imposte");
                B58.setCellStyle(grassetto);
                C58.setCellFormula("C28+C30+C53");
                C58.setCellStyle(grassetto);


// Creating row 59
                XSSFRow row59 = sheet.createRow(58);
                XSSFCell A59 = row59.createCell(0);
                XSSFCell B59 = row59.createCell(1);
                XSSFCell C59 = row59.createCell(2);
                A59.setCellValue("22)");
                B59.setCellValue("Imposte sul reddito dell'esercizio");
                r59=findViewById(R.id.edittext_imposte_reddito_esercizio);
                if(r59.getText().toString().length()!=0){
                    double riga59= Double.parseDouble(r59.getText().toString());
                    C59.setCellValue(riga59);
                }

// Creating row 60
                XSSFRow row60 = sheet.createRow(59);
                XSSFCell A60 = row60.createCell(0);
                XSSFCell B60 = row60.createCell(1);
                XSSFCell C60 = row60.createCell(2);
                A60.setCellValue("26)");
                A60.setCellStyle(grassetto);
                B60.setCellValue("Utile (perdita) dell'esercizio");
                B60.setCellStyle(grassetto);
                C60.setCellFormula("C58-C59");
                C60.setCellStyle(grassetto);

// Creating row 61
                XSSFRow row61 = sheet.createRow(60);
                XSSFCell A61 = row61.createCell(0);
                XSSFCell B61 = row61.createCell(1);
                XSSFCell C61 = row61.createCell(2);

                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                evaluator.evaluateAll();

                RegionUtil.setBorderBottom(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("A60:C60"), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("A3:C3"), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("A3:A60"), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("B3:B60"), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("C3:C60"), sheet);
                RegionUtil.setBorderLeft(BorderStyle.MEDIUM,
                        CellRangeAddress.valueOf("A3:A60"), sheet);

                for(int i=4; i<=60; i++) {
                    RegionUtil.setBorderTop(BorderStyle.THIN, CellRangeAddress.valueOf("A"+i+":C"+i), sheet);
                }


                sheet.setColumnWidth(1, 20000);
                // Salvataggio del file
                fileSaver.execute(workbook);

                }
});
        }
        public double getRicavi(){
        return ricavi;
        }

}