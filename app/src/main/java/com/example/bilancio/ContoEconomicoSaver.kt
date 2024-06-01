import android.content.Context
import android.os.AsyncTask
import android.os.Environment
import android.widget.Toast
import org.apache.poi.ss.usermodel.Workbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

class ContoEconomicoSaver(private val context: Context) : AsyncTask<Workbook, Void, Boolean>() {

    override fun doInBackground(vararg workbooks: Workbook): Boolean {
        return try {
            val folder =
                File(context.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO")
            if (!folder.exists()) {
                folder.mkdirs()
            }
            var file = File(folder, "contoeconomico.xlsx")
            var count = 1
            while (file.exists()) {
                val filename = "contoeconomico($count).xlsx"
                file = File(folder, filename)
                count++
            }
            FileOutputStream(file).use { outputStream -> workbooks[0].write(outputStream) }
            true
        } catch (e: IOException) {
            e.printStackTrace()
            false
        }
    }

    override fun onPostExecute(result: Boolean) {
        if (result) {
            Toast.makeText(context, "CONTO ECONOMICO SALVATO CORRETTAMENTE", Toast.LENGTH_SHORT)
                .show()
        } else {
            Toast.makeText(context, "ERRORE durante il salvataggio del file", Toast.LENGTH_SHORT)
                .show()
        }
    }
}
