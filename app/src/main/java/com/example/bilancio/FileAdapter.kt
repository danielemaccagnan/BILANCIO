package com.example.bilancio

import android.content.Context
import android.content.Intent
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import android.widget.Toast
import androidx.core.content.FileProvider
import androidx.recyclerview.widget.RecyclerView
import com.example.bilancio.FileAdapter.FileViewHolder
import java.io.File

class FileAdapter(private val context: Context, private val mFiles: ArrayList<File>) :
    RecyclerView.Adapter<FileViewHolder>() {
    private val selectedItems: MutableSet<Int> = HashSet()
    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): FileViewHolder {
        val itemView =
            LayoutInflater.from(parent.context).inflate(R.layout.file_item, parent, false)
        return FileViewHolder(itemView)
    }

    override fun onBindViewHolder(holder: FileViewHolder, position: Int) {
        val currentFile = mFiles[position]
        holder.textViewFileName.text = currentFile.name
        holder.itemView.setBackgroundColor(
            if (selectedItems.contains(position)) context.resources.getColor(
                android.R.color.holo_blue_light
            ) else context.resources.getColor(android.R.color.transparent)
        )
        holder.itemView.setOnClickListener { v: View? -> openFile(currentFile) }
        holder.itemView.setOnLongClickListener { v: View ->
            if (selectedItems.contains(position)) {
                selectedItems.remove(position)
                v.setBackgroundColor(context.resources.getColor(android.R.color.transparent))
            } else {
                selectedItems.add(position)
                v.setBackgroundColor(context.resources.getColor(android.R.color.holo_blue_light))
            }
            true
        }
    }

    override fun getItemCount(): Int {
        return mFiles.size
    }

    val selectedFiles: ArrayList<File>
        get() {
            val selectedFiles = ArrayList<File>()
            for (position in selectedItems) {
                selectedFiles.add(mFiles[position])
            }
            return selectedFiles
        }

    fun removeSelectedItems() {
        if (selectedItems.isEmpty()) {
            Toast.makeText(context, "Nessun file selezionato.", Toast.LENGTH_SHORT).show()
            return
        }
        val toRemove = ArrayList<File>()
        for (position in selectedItems) {
            val file = mFiles[position]
            if (file.delete()) {
                toRemove.add(file)
                Toast.makeText(context, file.name + " eliminato.", Toast.LENGTH_SHORT).show()
            } else {
                Toast.makeText(
                    context,
                    "Errore durante l'eliminazione di " + file.name,
                    Toast.LENGTH_SHORT
                ).show()
            }
        }
        mFiles.removeAll(toRemove)
        selectedItems.clear()
        notifyDataSetChanged()
    }

    private fun openFile(file: File) {
        val uri = FileProvider.getUriForFile(
            context,
            context.applicationContext.packageName + ".provider",
            file
        )
        val intent = Intent(Intent.ACTION_VIEW)
        intent.setDataAndType(
            uri,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        if (intent.resolveActivity(context.packageManager) != null) {
            context.startActivity(intent)
        } else {
            Toast.makeText(context, "Nessuna app trovata per aprire questo file", Toast.LENGTH_LONG)
                .show()
        }
    }

    class FileViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView) {
        var textViewFileName: TextView

        init {
            textViewFileName = itemView.findViewById(R.id.textViewFileName)
        }
    }
}
