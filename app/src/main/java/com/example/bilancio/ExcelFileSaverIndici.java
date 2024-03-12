package com.example.bilancio;

import android.content.Context;
import android.os.AsyncTask;
import android.os.Environment;
import android.widget.Toast;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFileSaverIndici  extends AsyncTask<Workbook, Void, Boolean> {

    private Context context;

    public ExcelFileSaverIndici(Context context) {
        this.context = context;
    }

    @Override
    protected Boolean doInBackground(Workbook... workbooks) {
        try {
            File folder = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), "BILANCIO");
            folder.mkdirs();
            File file = new File(folder, "indici.xlsx");
            int count = 1;
            while (file.exists()) {
                String filename = "indici(" + count + ").xlsx";
                file = new File(folder, filename);
                count++;
            }
            FileOutputStream outputStream = new FileOutputStream(file);
            workbooks[0].write(outputStream);
            outputStream.flush();
            outputStream.close();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }


    @Override
    protected void onPostExecute(Boolean result) {
        if (result) {
            Toast.makeText(context, "INDICI SALVATI CORRETTAMENTE", Toast.LENGTH_SHORT).show();
        } else {
            Toast.makeText(context, "ERRORE durante il salvataggio del file", Toast.LENGTH_SHORT).show();
        }
    }
}
