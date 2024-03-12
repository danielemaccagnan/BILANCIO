package com.example.bilancio;

import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.core.content.FileProvider;
import androidx.recyclerview.widget.RecyclerView;

import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

public class FileAdapter extends RecyclerView.Adapter<FileAdapter.FileViewHolder> {

    private ArrayList<File> mFiles;
    private Context context;
    private Set<Integer> selectedItems = new HashSet<>();
    private boolean isSelectionMode = false;

    public FileAdapter(Context context, ArrayList<File> files) {
        this.context = context;
        this.mFiles = files;
    }

    @NonNull
    @Override
    public FileViewHolder onCreateViewHolder(@NonNull ViewGroup parent, int viewType) {
        View itemView = LayoutInflater.from(parent.getContext()).inflate(R.layout.file_item, parent, false);
        return new FileViewHolder(itemView);
    }

    @Override
    public void onBindViewHolder(@NonNull FileViewHolder holder, int position) {
        File currentFile = mFiles.get(position);
        holder.textViewFileName.setText(currentFile.getName());
        // L'icona di Excel è già impostata nel layout XML, quindi non serve fare altro qui per visualizzarla

        holder.itemView.setOnClickListener(v -> {
            if (isSelectionMode) {
                toggleItemSelection(position, holder.itemView);
            } else {
                openFile(currentFile);
            }
        });

        holder.itemView.setOnLongClickListener(v -> {
            if (!isSelectionMode) {
                isSelectionMode = true;
                toggleItemSelection(position, holder.itemView);
            } else {
                clearSelection();
            }
            return true;
        });

        holder.itemView.setBackgroundColor(selectedItems.contains(position) ? context.getResources().getColor(android.R.color.holo_blue_light) : context.getResources().getColor(android.R.color.transparent));
    }

    private void openFile(File file) {
        Uri uri = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName() + ".provider", file);
        Intent intent = new Intent(Intent.ACTION_VIEW);
        intent.setDataAndType(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        if (intent.resolveActivity(context.getPackageManager()) != null) {
            context.startActivity(intent);
        } else {
            Toast.makeText(context, "Nessuna app trovata per aprire questo file", Toast.LENGTH_LONG).show();
        }
    }

    private void toggleItemSelection(int position, View itemView) {
        if (selectedItems.contains(position)) {
            selectedItems.remove(position);
            itemView.setBackgroundColor(context.getResources().getColor(android.R.color.transparent));
        } else {
            selectedItems.add(position);
            itemView.setBackgroundColor(context.getResources().getColor(android.R.color.holo_blue_light));
        }
    }

    @Override
    public int getItemCount() {
        return mFiles.size();
    }

    public void removeSelectedItems() {
        if (selectedItems.isEmpty()) {
            Toast.makeText(context, "Nessun file selezionato.", Toast.LENGTH_SHORT).show();
            return;
        }

        Set<File> filesToDelete = new HashSet<>();
        for (Integer position : selectedItems) {
            filesToDelete.add(mFiles.get(position));
        }

        for (File file : filesToDelete) {
            if (file.delete()) {
                mFiles.remove(file);
                Toast.makeText(context, file.getName() + " eliminato.", Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(context, "Errore durante l'eliminazione di " + file.getName(), Toast.LENGTH_SHORT).show();
            }
        }
        clearSelection();
        notifyDataSetChanged(); // Aggiorna l'adapter per riflettere le modifiche.
    }


    private void clearSelection() {
        selectedItems.clear();
        isSelectionMode = false;
        notifyDataSetChanged();
    }

    static class FileViewHolder extends RecyclerView.ViewHolder {
        TextView textViewFileName;
        ImageView imageViewExcelIcon; // Aggiunto riferimento all'ImageView

        public FileViewHolder(View itemView) {
            super(itemView);
            textViewFileName = itemView.findViewById(R.id.textViewFileName);
            imageViewExcelIcon = itemView.findViewById(R.id.imageViewExcelIcon); // Collegamento all'ImageView nell'XML
        }
    }
}