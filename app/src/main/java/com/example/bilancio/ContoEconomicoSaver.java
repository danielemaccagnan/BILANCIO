package com.example.bilancio;

import android.content.Context;
import android.os.AsyncTask;
import android.os.Environment;
import android.widget.Toast;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ContoEconomicoSaver extends AsyncTask<Workbook, Void, Boolean> {

    private Context context;

    public ContoEconomicoSaver(Context context) {
        this.context = context;
    }

    @Override
    protected Boolean doInBackground(Workbook... workbooks) {
        try {
            // Usa il contesto per chiamare getExternalFilesDir
            File folder = new File(context.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO");
            if (!folder.exists()) {
                folder.mkdirs(); // Crea la directory se non esiste
            }
            File file = new File(folder, "contoeconomico.xlsx");
            int count = 1;
            while (file.exists()) {
                String filename = "contoeconomico(" + count + ").xlsx";
                file = new File(folder, filename);
                count++;
            }
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                workbooks[0].write(outputStream);
            }
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }



    @Override
    protected void onPostExecute(Boolean result) {
        if (result) {
            Toast.makeText(context, "CONTO ECONOMICO SALVATO CORRETTAMENTE", Toast.LENGTH_SHORT).show();
        } else {
            Toast.makeText(context, "ERRORE durante il salvataggio del file", Toast.LENGTH_SHORT).show();
        }
    }
}
