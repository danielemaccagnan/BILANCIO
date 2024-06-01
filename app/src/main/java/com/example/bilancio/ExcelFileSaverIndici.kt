package com.example.bilancio

import android.content.Context
import android.os.AsyncTask
import android.os.Environment
import android.widget.Toast
import org.apache.poi.ss.usermodel.Workbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

class ExcelFileSaverIndici(private val context: Context)  {
    protected  fun doInBackground(vararg workbooks: Workbook): Boolean {
        return try {
            val folder = File(
                Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS),
                "BILANCIO"
            )
            folder.mkdirs()
            var file = File(folder, "indici.xlsx")
            var count = 1
            while (file.exists()) {
                val filename = "indici($count).xlsx"
                file = File(folder, filename)
                count++
            }
            val outputStream = FileOutputStream(file)
            workbooks[0].write(outputStream)
            outputStream.flush()
            outputStream.close()
            true
        } catch (e: IOException) {
            e.printStackTrace()
            false
        }
    }

     fun onPostExecute(result: Boolean) {
        if (result) {
            Toast.makeText(context, "INDICI SALVATI CORRETTAMENTE", Toast.LENGTH_SHORT).show()
        } else {
            Toast.makeText(context, "ERRORE durante il salvataggio del file", Toast.LENGTH_SHORT)
                .show()
        }
    }
}
