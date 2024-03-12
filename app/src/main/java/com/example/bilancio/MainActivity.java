package com.example.bilancio;

import android.content.Intent;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.widget.Button;

import androidx.appcompat.app.AppCompatActivity;
import androidx.preference.PreferenceManager;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // Impostazioni dei pulsanti
        Button statopatrimoniale = findViewById(R.id.button_statopatrimoniale);
        Button contoeconomico = findViewById(R.id.button_conto_economico);
        Button cambiaindici = findViewById(R.id.button_indici);
        Button file_salvati = findViewById(R.id.button_filesalvati);
        Button menu = findViewById(R.id.menu_main);

        // Imposta il tema di default se non è già stato fatto
        SharedPreferences prefs = PreferenceManager.getDefaultSharedPreferences(this);
        if (!prefs.contains("theme")) {
            prefs.edit().putString("theme", "system_default").apply();
        }

        // Listener per i pulsanti che avviano le altre Activity
        menu.setOnClickListener(view -> startActivity(new Intent(MainActivity.this, SettingsActivity.class)));
        file_salvati.setOnClickListener(v -> startActivity(new Intent(MainActivity.this, com.example.bilancio.filesalvati.class)));
        cambiaindici.setOnClickListener(view -> startActivity(new Intent(MainActivity.this, indici.class)));
        statopatrimoniale.setOnClickListener(v -> startActivity(new Intent(MainActivity.this, statopatrimoniale.class)));
        contoeconomico.setOnClickListener(v -> startActivity(new Intent(MainActivity.this, contoeconomico.class)));
    }
}
