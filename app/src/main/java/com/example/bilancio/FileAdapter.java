package com.example.bilancio;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;

import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

public class FileAdapter extends RecyclerView.Adapter<FileAdapter.FileViewHolder> {
    private ArrayList<File> mFiles;
    private Context context;
    private Set<Integer> selectedItems = new HashSet<>();

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

        holder.itemView.setBackgroundColor(selectedItems.contains(position) ? context.getResources().getColor(android.R.color.holo_blue_light) : context.getResources().getColor(android.R.color.transparent));

        holder.itemView.setOnClickListener(v -> {
            if (selectedItems.contains(position)) {
                selectedItems.remove(position);
            } else {
                selectedItems.add(position);
            }
            notifyItemChanged(position);
        });
    }

    @Override
    public int getItemCount() {
        return mFiles.size();
    }

    public ArrayList<File> getSelectedFiles() {
        ArrayList<File> selectedFiles = new ArrayList<>();
        for (Integer position : selectedItems) {
            selectedFiles.add(mFiles.get(position));
        }
        return selectedFiles;
    }

    public void removeSelectedItems() {
        if (selectedItems.isEmpty()) {
            Toast.makeText(context, "Nessun file selezionato.", Toast.LENGTH_SHORT).show();
            return;
        }

        ArrayList<File> toRemove = new ArrayList<>();
        for (Integer position : selectedItems) {
            File file = mFiles.get(position);
            if (file.delete()) {
                toRemove.add(file);
                Toast.makeText(context, file.getName() + " eliminato.", Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(context, "Errore durante l'eliminazione di " + file.getName(), Toast.LENGTH_SHORT).show();
            }
        }
        mFiles.removeAll(toRemove);
        selectedItems.clear();
        notifyDataSetChanged();
    }

    static class FileViewHolder extends RecyclerView.ViewHolder {
        TextView textViewFileName;

        public FileViewHolder(View itemView) {
            super(itemView);
            textViewFileName = itemView.findViewById(R.id.textViewFileName);
        }
    }
}
