package com.example.bilancio;

import android.content.Intent;
import android.os.Bundle;
import android.os.Environment;
import android.widget.Button;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
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

        // Setup del bottone per tornare alla MainActivity
        Button homeButton = findViewById(R.id.button_filesalvati);
        Button deletebutton = findViewById(R.id.delete_button);
        homeButton.setOnClickListener(view -> startActivity(new Intent(filesalvati.this, MainActivity.class)));
        deletebutton.setOnClickListener(view -> {
            if (adapter != null) {
                adapter.removeSelectedItems();
            }
        });

        // Impostazione del RecyclerView
        recyclerView = findViewById(R.id.recyclerViewFiles);
        recyclerView.setLayoutManager(new LinearLayoutManager(this));

        // Popolamento del RecyclerView con i file trovati
        populateRecyclerView();
    }

    private void populateRecyclerView() {
        // Ottiene il riferimento alla directory dei documenti privati dell'app sotto "BILANCIO"
        File directory = new File(getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO");

        // Verifica se la directory esiste
        if (!directory.exists()) {
            // Se la directory non esiste, mostra un Toast e termina il metodo
            Toast.makeText(this, "La directory BILANCIO non esiste.", Toast.LENGTH_LONG).show();
            return;
        }

        // Se la directory esiste, procedi con la ricerca dei file Excel
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
                    // Ricerca ricorsiva in sottodirectory per trovare file Excel
                    fileList.addAll(findExcelFiles(file));
                } else {
                    // Aggiunge alla lista solo i file che terminano con .xls o .xlsx
                    if (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) {
                        fileList.add(file);
                    }
                }
            }
        }
        return fileList;
    }
}
