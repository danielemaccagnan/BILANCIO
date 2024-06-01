package com.example.bilancio

import android.content.Intent
import android.os.Bundle
import android.view.View
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.preference.PreferenceManager

class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // Impostazioni dei pulsanti
        val statopatrimonialee = findViewById<TextView>(R.id.button_statopatrimoniale)
        val contoeconomicoo = findViewById<TextView>(R.id.button_conto_economico)
        val cambiaindici = findViewById<TextView>(R.id.button_indici)
        val file_salvati = findViewById<TextView>(R.id.button_filesalvati)
        val menu = findViewById<TextView>(R.id.menu_main)

        // Imposta il tema di default se non è già stato fatto
        val prefs = PreferenceManager.getDefaultSharedPreferences(this)
        if (!prefs.contains("theme")) {
            prefs.edit().putString("theme", "system_default").apply()
        }

        // Listener per i pulsanti che avviano le altre Activity
        menu.setOnClickListener { view: View? ->
            startActivity(
                Intent(
                    this@MainActivity,
                    SettingsActivity::class.java
                )
            )
        }
        file_salvati.setOnClickListener { v: View? ->
            startActivity(
                Intent(
                    this@MainActivity,
                    filesalvati::class.java
                )
            )
        }
        cambiaindici.setOnClickListener { view: View? ->
            startActivity(
                Intent(
                    this@MainActivity,
                    indici::class.java
                )
            )
        }
        statopatrimonialee.setOnClickListener { v: View? ->
            startActivity(
                Intent(
                    this@MainActivity,
                    statopatrimoniale::class.java
                )
            )
        }
        contoeconomicoo.setOnClickListener { v: View? ->
            startActivity(
                Intent(
                    this@MainActivity,
                    contoeconomico::class.java
                )
            )
        }
    }
}