package com.example.bilancio;

import android.content.Context;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;

import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class indici extends AppCompatActivity {
    Context context = this;

    // Dichiarazione degli EditText come membri della classe
    public EditText capitale_di_terzi_ind_globale, patrimonio_netto_ind_globale, banca_ind_strutturale,
            quotamutuipassivi_ind_strutturale, mutuipassivi_ind_strutturale, patrimonio_netto_ind_strutturale, pas_brevetermine_ind_esigibilità, patrimonio_netto_ind_esigibilità,
            attivo_fisso_netto_ind_rigidità_attivo, attivo_totale_netto_ind_rigidità_attivo, attivo_a_breve_ind_elasticità_attivo,
            attivo_totale_netto_ind_elasticità_attivo, patrimonio_netto_ind_grado_di_copertura_1,
            attivo_fisso_netto_ind_grado_di_copertura_1, patrimonio_netto_ind_grado_di_copertura_2,
            passivo_medio_lungo_termine_ind_grado_di_copertura_2, attivo_fisso_netto_ind_grado_di_copertura_2,
            fondo_ammortamento_ind_grado_di_ammortamento, immobilizzazioni_materiali_ind_grado_di_ammortamento,
            roa_ind_roi, attivo_totale_netto_ind_roi, roa_ind_ros, ricavi_netti_vendita_ind_ros,
            reddito_netto_utile_vendita_ind_roe, patrimonio_netto_ind_roe, ricavi_netti_vendita_ind_turnover,
            attivo_totale_netto_ind_turnover, cassa_ind_liquidità_1, banca_ind_liquidità_1, crediti_ind_liquidità_1,
            passività_ind_liquidità_1, attività_ind_liquidità_2, passività_ind_liquidità_2;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_indici);

        // Inizializzazione degli EditText
        capitale_di_terzi_ind_globale = findViewById(R.id.capitale_di_terzi_ind_globale);
        patrimonio_netto_ind_globale = findViewById(R.id.patrimonio_netto_ind_globale);
        banca_ind_strutturale = findViewById(R.id.banca_ind_strutturale);
        quotamutuipassivi_ind_strutturale = findViewById(R.id.quotamutuipassivi_ind_strutturale);
        mutuipassivi_ind_strutturale = findViewById(R.id.mutuipassivi_ind_strutturale);
        patrimonio_netto_ind_strutturale=findViewById(R.id.patrimonio_netto_ind_strutturale);
        pas_brevetermine_ind_esigibilità = findViewById(R.id.pas_brevetermine_ind_esigibilità);
        patrimonio_netto_ind_esigibilità = findViewById(R.id.patrimonio_netto_ind_esigibilità);
        attivo_fisso_netto_ind_rigidità_attivo = findViewById(R.id.attivo_fisso_netto_ind_rigidità_attivo);
        attivo_totale_netto_ind_rigidità_attivo = findViewById(R.id.attivo_totale_netto_ind_rigidità_attivo);
        attivo_a_breve_ind_elasticità_attivo = findViewById(R.id.attivo_a_breve_ind_elasticità_attivo);
        attivo_totale_netto_ind_elasticità_attivo = findViewById(R.id.attivo_totale_netto_ind_elasticità_attivo);
        patrimonio_netto_ind_grado_di_copertura_1 = findViewById(R.id.patrimonio_netto_ind_grado_di_copertura_1);
        attivo_fisso_netto_ind_grado_di_copertura_1 = findViewById(R.id.attivo_fisso_netto_ind_grado_di_copertura_1);
        patrimonio_netto_ind_grado_di_copertura_2 = findViewById(R.id.patrimonio_netto_ind_grado_di_copertura_2);
        passivo_medio_lungo_termine_ind_grado_di_copertura_2 = findViewById(R.id.passivo_medio_lungo_termine_ind_grado_di_copertura_2);
        attivo_fisso_netto_ind_grado_di_copertura_2 = findViewById(R.id.attivo_fisso_netto_ind_grado_di_copertura_2);
        fondo_ammortamento_ind_grado_di_ammortamento = findViewById(R.id.fondo_ammortamento_ind_grado_di_ammortamento);
        immobilizzazioni_materiali_ind_grado_di_ammortamento = findViewById(R.id.immobilizzazioni_materiali_ind_grado_di_ammortamento);
        roa_ind_roi = findViewById(R.id.roa_ind_roi);
        attivo_totale_netto_ind_roi = findViewById(R.id.attivo_totale_netto_ind_roi);
        roa_ind_ros = findViewById(R.id.roa_ind_ros);
        ricavi_netti_vendita_ind_ros = findViewById(R.id.ricavi_netti_vendita_ind_ros);
        reddito_netto_utile_vendita_ind_roe = findViewById(R.id.reddito_netto_utile_vendita_ind_roe);
        patrimonio_netto_ind_roe = findViewById(R.id.patrimonio_netto_ind_roe);
        ricavi_netti_vendita_ind_turnover = findViewById(R.id.ricavi_netti_vendita_ind_turnover);
        attivo_totale_netto_ind_turnover = findViewById(R.id.attivo_totale_netto_ind_turnover);
        cassa_ind_liquidità_1 = findViewById(R.id.cassa_ind_liquidità_1);
        banca_ind_liquidità_1 = findViewById(R.id.banca_ind_liquidità_1);
        crediti_ind_liquidità_1 = findViewById(R.id.crediti_ind_liquidità_1);
        passività_ind_liquidità_1 = findViewById(R.id.passività_ind_liquidità_1);
        attività_ind_liquidità_2 = findViewById(R.id.attività_ind_liquidità_2);
        passività_ind_liquidità_2 = findViewById(R.id.passività_ind_liquidità_2);


        Button salvaa = findViewById(R.id.salvaindici);
        salvaa.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                // Crea una nuova istanza di IndiciDiBilancioSaver
                IndiciDiBilancioSaver fileSaver = new IndiciDiBilancioSaver(context);

                // Crea un nuovo workbook e imposta i valori
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Indici");

                // Ciclo per creare le righe con i valori
                for (int i = 0; i < 35; i++) {
                    XSSFRow row = sheet.createRow(i);
                    // Creazione delle 5 colonne per ogni riga
                    for (int j = 0; j < 4; j++) {
                        XSSFCell cell = row.createCell(j);
                        // Imposta i valori in base alle condizioni richieste
                        if (j == 0) {
                            cell.setCellValue(""); // Colonna A
                        } else if (j == 1) {
                            cell.setCellValue(""); // Colonna B
                        } else if (j == 2) {
                            cell.setCellValue(""); // Colonna C
                        } else if (j==3){
                            cell.setCellValue(""); // Colonna D
                        }
                    }
                }


                // Ottieni l'oggetto CellStyle per le celle unite
                XSSFCellStyle mergedCellStyle = workbook.createCellStyle();
                mergedCellStyle.setBorderBottom(BorderStyle.THIN); // Bordo inferiore
                mergedCellStyle.setBorderTop(BorderStyle.THIN); // Bordo superiore
                mergedCellStyle.setBorderLeft(BorderStyle.THIN); // Bordo sinistro
                mergedCellStyle.setBorderRight(BorderStyle.THIN); // Bordo destro

                // Stile orizzontale
                XSSFCellStyle horizontalstyle = workbook.createCellStyle();
                // Imposta l'allineamento del testo al centro
                horizontalstyle.setAlignment(HorizontalAlignment.CENTER);

                // Bordo piccolo
                RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, 3), sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(1, 1, 0, 1), sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(2, 2, 0, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(2, 2, 3, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(2, 2, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(1, 1, 0, 0), sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(3, 3, 0, 0), sheet);

                // Bordo medio destro
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(4, 5, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(8, 9, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(12, 13, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(16, 17, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(20, 21, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(24, 25, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(28, 29, 3, 3), sheet);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(32, 33, 3, 3), sheet);

                //Bordo medio alto
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(7, 7, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(11, 11, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(15, 15, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(19, 19, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(23, 23, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(27, 27, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(31, 31, 0, 3), sheet);
                RegionUtil.setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(35, 35, 0, 3), sheet);

                //Bordo medio basso
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(3, 3, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(7, 7, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(11, 11, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(15, 15, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(19, 19, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(23, 23, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(27, 27, 0, 3), sheet);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(31, 31, 0, 3), sheet);

                // Imposta la larghezza delle colonne
                sheet.setColumnWidth(0, 5000);
                sheet.setColumnWidth(1, 5000);
                sheet.setColumnWidth(2, 5000);
                sheet.setColumnWidth(3, 5000);














                // Unisci celle
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
                sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 2));
                sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(16, 16, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(20, 20, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(24, 24, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(28, 28, 0, 1));
                sheet.addMergedRegion(new CellRangeAddress(32, 32, 0, 1));

                // Stile per il testo in grassetto e colore rosso
                XSSFCellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                font.setColor(IndexedColors.RED.getIndex());
                style.setFont(font);


                XSSFRow rowA1 = sheet.getRow(0);
                XSSFCell cellA1 = rowA1.createCell(0);
                cellA1.setCellValue("INDICI DI BILANCIO - SOLIDITÀ");


                XSSFRow rowA3 = sheet.getRow(2);
                XSSFCell cellA3 = rowA3.createCell(0);
                cellA3.setCellValue("INDICI DI SOLIDITÀ");

                XSSFRow rowA5 = sheet.getRow(4);
                XSSFCell cellA5 = rowA5.createCell(0);
                cellA5.setCellStyle(style);
                cellA5.setCellValue("INDEBITAMENTO GLOBALE");

                XSSFRow rowC5 = sheet.getRow(4);
                XSSFCell cellC5 = rowC5.createCell(2);
                cellC5.setCellValue("CAPITALE DI TERZI");

                XSSFRow rowC6 = sheet.getRow(5);
                XSSFCell cellC6 = rowC6.createCell(2);
                cellC6.setCellValue("PATRIMONIO NETTO");

                XSSFRow rowA7 = sheet.getRow(6);
                XSSFCell cellA7 = rowA7.createCell(0);
                cellA7.setCellValue("TOTALE");

                XSSFRow rowD7 = sheet.getRow(6);
                XSSFCell cellD7 = rowD7.createCell(3);

                if(capitale_di_terzi_ind_globale.getText().toString().length()!=0 && patrimonio_netto_ind_globale.getText().toString().length()!=0){
                    double capitalediterziindglobale=Double.parseDouble(capitale_di_terzi_ind_globale.getText().toString());
                    double patrimonionettoindglobale= Double.parseDouble(patrimonio_netto_ind_globale.getText().toString());
                    cellD7.setCellValue(capitalediterziindglobale/patrimonionettoindglobale);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(6, 6, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(6, 6, 3, 3), sheet);
                }

                XSSFRow rowA9 = sheet.getRow(8);
                XSSFCell cellA9 = rowA9.createCell(0);
                cellA9.setCellValue("INDEBITAMENTO STRUTTURALE");

                XSSFRow rowC9 = sheet.getRow(8);
                XSSFCell cellC9 = rowC9.createCell(2);
                cellA9.setCellStyle(style);
                cellC9.setCellValue("PASSIVITÀ FINANZIARIE");

                XSSFRow rowC10 = sheet.getRow(9);
                XSSFCell cellC10 = rowC10.createCell(2);
                cellC10.setCellValue("PATRIMONIO NETTO");

                XSSFRow rowA11 = sheet.getRow(10);
                XSSFCell cellA11 = rowA11.createCell(0);
                cellA11.setCellValue("TOTALE");

                XSSFRow rowD11 = sheet.getRow(10);
                XSSFCell cellD11 = rowD11.createCell(3);


                if(banca_ind_strutturale.getText().toString().length()!=0 || quotamutuipassivi_ind_strutturale.getText().toString().length()!=0 ||mutuipassivi_ind_strutturale.getText().toString().length()!=0 && patrimonio_netto_ind_strutturale.getText().toString().length()!=0){
                    double bancaccpassivostrutturale=Double.parseDouble(banca_ind_strutturale.getText().toString());
                    double quotamutuipassivistrutturale= Double.parseDouble(quotamutuipassivi_ind_strutturale.getText().toString());
                    double mutuipassiviindstrutturale = Double.parseDouble(mutuipassivi_ind_strutturale.getText().toString());
                    double patrimonionettoindstrutturale = Double.parseDouble(patrimonio_netto_ind_strutturale.getText().toString());
                    cellD11.setCellValue((bancaccpassivostrutturale + quotamutuipassivistrutturale + mutuipassiviindstrutturale) /patrimonionettoindstrutturale);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(10, 10, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(10, 10, 3, 3), sheet);
                }

                XSSFRow rowA13 = sheet.getRow(12);
                XSSFCell cellA13 = rowA13.createCell(0);
                cellA13.setCellStyle(style);
                cellA13.setCellValue("ESIGIBILITÀ DEL PASSIVO");

                XSSFRow rowC13 = sheet.getRow(12);
                XSSFCell cellC13 = rowC13.createCell(2);
                cellC13.setCellValue("PASSIVITÀ A BREVE TERMINE");

                XSSFRow rowC14 = sheet.getRow(13);
                XSSFCell cellC14 = rowC14.createCell(2);
                cellC14.setCellValue("CAPITALE DI TERZI");

                XSSFRow rowA15 = sheet.getRow(14);
                XSSFCell cellA15 = rowA15.createCell(0);
                cellA15.setCellValue("TOTALE");

                XSSFRow rowD15 = sheet.getRow(14);
                XSSFCell cellD15 = rowD15.createCell(3);

                if(pas_brevetermine_ind_esigibilità.getText().toString().length()!=0 && patrimonio_netto_ind_esigibilità.getText().toString().length()!=0){
                    double passivitaabreveindpassivo=Double.parseDouble(pas_brevetermine_ind_esigibilità.getText().toString());
                    double patrimonionettoindpassivo= Double.parseDouble(patrimonio_netto_ind_esigibilità.getText().toString());
                    cellD15.setCellValue(passivitaabreveindpassivo/patrimonionettoindpassivo);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(14, 14, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(14, 14, 3, 3), sheet);
                }

                XSSFRow rowA17 = sheet.getRow(16);
                XSSFCell cellA17 = rowA17.createCell(0);
                cellA17.setCellStyle(style);
                cellA17.setCellValue("GRADO DI RIGIDITÀ DELL'ATTIVO");

                XSSFRow rowC17 = sheet.getRow(16);
                XSSFCell cellC17 = rowC17.createCell(2);
                cellC17.setCellValue("ATTIVO FISSO NETTO");

                XSSFRow rowC18 = sheet.getRow(17);
                XSSFCell cellC18 = rowC18.createCell(2);
                cellC18.setCellValue("ATTIVO TOTALE NETTO");

                XSSFRow rowA19 = sheet.getRow(18);
                XSSFCell cellA19 = rowA19.createCell(0);
                cellA19.setCellValue("TOTALE");

                XSSFRow rowD19 = sheet.getRow(18);
                XSSFCell cellD19 = rowD19.createCell(3);

                if(attivo_fisso_netto_ind_rigidità_attivo.getText().toString().length()!=0 && attivo_totale_netto_ind_rigidità_attivo.getText().toString().length()!=0){
                    double attivofissonettorigidità=Double.parseDouble(attivo_fisso_netto_ind_rigidità_attivo.getText().toString());
                    double attivototalenettorigidità= Double.parseDouble(attivo_totale_netto_ind_rigidità_attivo.getText().toString());
                    cellD19.setCellValue(attivofissonettorigidità/attivototalenettorigidità);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(18, 18, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(18, 18, 3, 3), sheet);
                }

                XSSFRow rowA21 = sheet.getRow(20);
                XSSFCell cellA21 = rowA21.createCell(0);
                cellA21.setCellStyle(style);
                cellA21.setCellValue("GRADO DI ELASTICITÀ DELL'ATTIVO");

                XSSFRow rowC21 = sheet.getRow(20);
                XSSFCell cellC21 = rowC21.createCell(2);
                cellC21.setCellValue("ATTIVO A BREVE");

                XSSFRow rowC22 = sheet.getRow(21);
                XSSFCell cellC22 = rowC22.createCell(2);
                cellC22.setCellValue("ATTIVO TOTALE NETTO");

                XSSFRow rowA23 = sheet.getRow(22);
                XSSFCell cellA23 = rowA23.createCell(0);
                cellA23.setCellValue("TOTALE");

                XSSFRow rowD23 = sheet.getRow(22);
                XSSFCell cellD23 = rowD23.createCell(3);

                if(attivo_a_breve_ind_elasticità_attivo.getText().toString().length()!=0 && attivo_totale_netto_ind_elasticità_attivo.getText().toString().length()!=0){
                    double attivoabreveelasticità=Double.parseDouble(attivo_a_breve_ind_elasticità_attivo.getText().toString());
                    double attivototalenettoelasticità= Double.parseDouble(attivo_totale_netto_ind_elasticità_attivo.getText().toString());
                    cellD23.setCellValue(attivoabreveelasticità/attivototalenettoelasticità);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(22, 22, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(22, 22, 3, 3), sheet);
                }

                XSSFRow rowA25 = sheet.getRow(24);
                XSSFCell cellA25 = rowA25.createCell(0);
                cellA25.setCellStyle(style);
                cellA25.setCellValue("GRADO DI COPERTURA DELL'ATTIVO (I) LIVELLO");

                XSSFRow rowC25 = sheet.getRow(24);
                XSSFCell cellC25 = rowC25.createCell(2);
                cellC25.setCellValue("PATRIMONIO NETTO");

                XSSFRow rowC26 = sheet.getRow(25);
                XSSFCell cellC26 = rowC26.createCell(2);
                cellC26.setCellValue("ATTIVO FISSO NETTO");

                XSSFRow rowA27 = sheet.getRow(26);
                XSSFCell cellA27 = rowA27.createCell(0);
                cellA27.setCellValue("TOTALE");

                XSSFRow rowD27 = sheet.getRow(26);
                XSSFCell cellD27 = rowD27.createCell(3);

                if(patrimonio_netto_ind_grado_di_copertura_1.getText().toString().length()!=0 && attivo_fisso_netto_ind_grado_di_copertura_1.getText().toString().length()!=0){
                    double patrimonionettocopertura=Double.parseDouble(patrimonio_netto_ind_grado_di_copertura_1.getText().toString());
                    double attivofissonettocopertura= Double.parseDouble(attivo_fisso_netto_ind_grado_di_copertura_1.getText().toString());
                    cellD27.setCellValue(patrimonionettocopertura/attivofissonettocopertura);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(26, 26, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(26, 26, 3, 3), sheet);
                }

                XSSFRow rowA29 = sheet.getRow(28);
                XSSFCell cellA29 = rowA29.createCell(0);
                cellA29.setCellStyle(style);
                cellA29.setCellValue("GRADO DI COPERTURA DELL'ATTIVO (II) LIVELLO");

                XSSFRow rowC29 = sheet.getRow(28);
                XSSFCell cellC29 = rowC29.createCell(2);
                cellC29.setCellValue("PATRIMONIO NETTO + \n" + "PASSIVO A MEDIO/LUNGO\n" + "TERMINE");

                XSSFRow rowC30 = sheet.getRow(29);
                XSSFCell cellC30 = rowC30.createCell(2);
                cellC30.setCellValue("ATTIVO FISSO NETTO");

                XSSFRow rowA31 = sheet.getRow(30);
                XSSFCell cellA31 = rowA31.createCell(0);
                cellA31.setCellValue("TOTALE");

                XSSFRow rowD31 = sheet.getRow(30);
                XSSFCell cellD31 = rowD31.createCell(3);

                if(patrimonio_netto_ind_grado_di_copertura_2.getText().toString().length()!=0 || passivo_medio_lungo_termine_ind_grado_di_copertura_2.getText().toString().length()!=0 && attivo_fisso_netto_ind_grado_di_copertura_2.getText().toString().length()!=0){
                    double patrimonionettocopertura2=Double.parseDouble(patrimonio_netto_ind_grado_di_copertura_2.getText().toString());
                    double passivomediolungotermine = Double.parseDouble(passivo_medio_lungo_termine_ind_grado_di_copertura_2.getText().toString());
                    double attivofissonettocopertura2= Double.parseDouble(attivo_fisso_netto_ind_grado_di_copertura_2.getText().toString());
                    cellD31.setCellValue((patrimonionettocopertura2 + passivomediolungotermine)/attivofissonettocopertura2);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(30, 30, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(30, 30, 3, 3), sheet);
                }

                XSSFRow rowA33 = sheet.getRow(32);
                XSSFCell cellA33 = rowA33.createCell(0);
                cellA33.setCellStyle(style);
                cellA33.setCellValue("GRADO DI AMMORTAMENTO");

                XSSFRow rowC33 = sheet.getRow(32);
                XSSFCell cellC33 = rowC33.createCell(2);
                cellC33.setCellValue("FONDO AMMORTAMENTO");

                XSSFRow rowC34 = sheet.getRow(33);
                XSSFCell cellC34 = rowC34.createCell(2);
                cellC34.setCellValue("IMMOBILIZZAZIONI MATERIALI");

                XSSFRow rowA35 = sheet.getRow(34);
                XSSFCell cellA35 = rowA35.createCell(0);
                cellA35.setCellValue("TOTALE");

                XSSFRow rowD35 = sheet.getRow(34);
                XSSFCell cellD35 = rowD35.createCell(3);

                if(fondo_ammortamento_ind_grado_di_ammortamento.getText().toString().length()!=0 && immobilizzazioni_materiali_ind_grado_di_ammortamento.getText().toString().length()!=0){
                    double fondoammortamentogrado=Double.parseDouble(fondo_ammortamento_ind_grado_di_ammortamento.getText().toString());
                    double immobilizzazionimateriali= Double.parseDouble(immobilizzazioni_materiali_ind_grado_di_ammortamento.getText().toString());
                    cellD35.setCellValue(fondoammortamentogrado/immobilizzazionimateriali);
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(34, 34, 3, 3), sheet);
                } else {
                    RegionUtil.setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(34, 34, 3, 3), sheet);
                }












                // Esegui il salvataggio del file
                fileSaver.execute(workbook);
            }
        });
    }
}
