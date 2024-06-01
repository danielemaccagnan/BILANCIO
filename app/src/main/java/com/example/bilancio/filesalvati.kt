package com.example.bilancio

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.view.View
import android.widget.Button
import android.widget.ImageView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.FileProvider
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import java.io.File

class filesalvati : AppCompatActivity() {
    private var recyclerView: RecyclerView? = null
    private var adapter: FileAdapter? = null
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.file_salvati)
        val homeButton = findViewById<Button>(R.id.button_filesalvati)
        val deleteButton = findViewById<Button>(R.id.delete_button)
        val imageViewShareIcon = findViewById<ImageView>(R.id.imageViewShareIcon)
        homeButton.setOnClickListener { view: View? -> finish() }
        deleteButton.setOnClickListener { view: View? -> adapter!!.removeSelectedItems() }
        imageViewShareIcon.setOnClickListener { v: View? ->
            val selectedFiles = adapter!!.selectedFiles
            shareFiles(selectedFiles)
        }
        recyclerView = findViewById(R.id.recyclerViewFiles)
        recyclerView!!.setLayoutManager(LinearLayoutManager(this))
        populateRecyclerView()
    }

    private fun populateRecyclerView() {
        val directory = File(getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO")
        if (!directory.exists()) {
            Toast.makeText(this, "La directory BILANCIO non esiste.", Toast.LENGTH_LONG).show()
            return
        }
        val fileList = findExcelFiles(directory)
        adapter = FileAdapter(this, fileList)
        recyclerView!!.adapter = adapter
    }

    private fun findExcelFiles(directory: File): ArrayList<File> {
        val fileList = ArrayList<File>()
        val files = directory.listFiles()
        if (files != null) {
            for (file in files) {
                if (file.isDirectory) {
                    fileList.addAll(findExcelFiles(file))
                } else if (file.name.endsWith(".xls") || file.name.endsWith(".xlsx")) {
                    fileList.add(file)
                }
            }
        }
        return fileList
    }

    private fun shareFiles(files: ArrayList<File>) {
        if (files.isEmpty()) {
            Toast.makeText(this, "Nessun file selezionato.", Toast.LENGTH_SHORT).show()
            return
        }
        val uris = ArrayList<Uri>()
        for (file in files) {
            val uri =
                FileProvider.getUriForFile(this, applicationContext.packageName + ".provider", file)
            uris.add(uri)
        }
        val shareIntent = Intent(Intent.ACTION_SEND_MULTIPLE)
        shareIntent.setType("application/*") // Modifica se necessario in base al tipo di file
        shareIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris)
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        startActivity(Intent.createChooser(shareIntent, "Condividi File"))
    }
}
