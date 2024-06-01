package com.example.bilancio

import android.content.Context
import android.os.AsyncTask
import android.os.Environment
import android.widget.Toast
import org.apache.poi.ss.usermodel.Workbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

class IndiciDiBilancioSaver(context: Context) : AsyncTask<Workbook?, Void?, Boolean>() {
    private val appContext: Context = context.applicationContext

    override fun doInBackground(vararg workbooks: Workbook?): Boolean {
        return try {
            val folder = File(appContext.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO")
            if (!folder.exists()) {
                folder.mkdirs()
            }
            var file = File(folder, "indicidibilancio.xlsx")
            var count = 1
            while (file.exists()) {
                val filename = "indicidibilancio($count).xlsx"
                file = File(folder, filename)
                count++
            }
            FileOutputStream(file).use { outputStream -> workbooks[0]?.write(outputStream) }
            true
        } catch (e: IOException) {
            e.printStackTrace()
            false
        }
    }

    override fun onPostExecute(result: Boolean) {
        val message = if (result) "INDICI DI BILANCIO SALVATI CORRETTAMENTE" else "ERRORE durante il salvataggio del file"
        Toast.makeText(appContext, message, Toast.LENGTH_SHORT).show()
    }
}
