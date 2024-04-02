package com.example.bilancio;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;

import java.io.File;
import java.util.ArrayList;

public class filesalvati extends AppCompatActivity {
    private RecyclerView recyclerView;
    private FileAdapter adapter;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.file_salvati);

        Button homeButton = findViewById(R.id.button_filesalvati);
        Button deleteButton = findViewById(R.id.delete_button);
        ImageView imageViewShareIcon = findViewById(R.id.imageViewShareIcon);

        homeButton.setOnClickListener(view -> finish());
        deleteButton.setOnClickListener(view -> {
            adapter.removeSelectedItems();
        });
        imageViewShareIcon.setOnClickListener(v -> {
            ArrayList<File> selectedFiles = adapter.getSelectedFiles();
            shareFiles(selectedFiles);
        });


        recyclerView = findViewById(R.id.recyclerViewFiles);
        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        populateRecyclerView();
    }

    private void populateRecyclerView() {
        File directory = new File(getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO");
        if (!directory.exists()) {
            Toast.makeText(this, "La directory BILANCIO non esiste.", Toast.LENGTH_LONG).show();
            return;
        }
        ArrayList<File> fileList = findExcelFiles(directory);
        adapter = new FileAdapter(this, fileList);
        recyclerView.setAdapter(adapter);
    }

    private ArrayList<File> findExcelFiles(File directory) {
        ArrayList<File> fileList = new ArrayList<>();
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    fileList.addAll(findExcelFiles(file));
                } else if (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) {
                    fileList.add(file);
                }
            }
        }
        return fileList;
    }

    private void shareFiles(ArrayList<File> files) {
        if (files.isEmpty()) {
            Toast.makeText(this, "Nessun file selezionato.", Toast.LENGTH_SHORT).show();
            return;
        }

        ArrayList<Uri> uris = new ArrayList<>();
        for (File file : files) {
            Uri uri = FileProvider.getUriForFile(this, getApplicationContext().getPackageName() + ".provider", file);
            uris.add(uri);
        }

        Intent shareIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
        shareIntent.setType("application/*"); // Modifica se necessario in base al tipo di file
        shareIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        startActivity(Intent.createChooser(shareIntent, "Condividi File"));


    }

}
