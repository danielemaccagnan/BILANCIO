package com.example.bilancio;

import android.content.Context;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class indici extends AppCompatActivity {
    Context context = this;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_indici);
        Button salvaa = findViewById(R.id.salvaindici);
        salvaa.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                // Crea una nuova istanza di IndiciDiBilancioSaver
                IndiciDiBilancioSaver fileSaver = new IndiciDiBilancioSaver(context);

                // Crea un nuovo workbook e imposta i valori
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Indici");
                XSSFRow row1 = sheet.createRow(0);
                XSSFCell A1 = row1.createCell(0);
                XSSFCell B1 = row1.createCell(1);
                XSSFCell C1 = row1.createCell(2);
                contoeconomico e = new contoeconomico();
                double x = e.getRicavi();
                A1.setCellValue(x);

                // Esegui il salvataggio del file
                fileSaver.execute(workbook);
            }
        });
    }
}
