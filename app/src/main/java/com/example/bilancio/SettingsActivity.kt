package com.example.bilancio

import android.os.Bundle
import androidx.appcompat.app.AppCompatActivity
import androidx.appcompat.app.AppCompatDelegate
import androidx.preference.PreferenceManager
import java.util.Locale

class SettingsActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        // Ottiene le SharedPreferences
        val preferences = PreferenceManager.getDefaultSharedPreferences(this)
        // Ottiene il tema selezionato, con "system" come valore di default
        val selectedTheme = preferences.getString("theme", "system")

        // Applica il tema in base alla selezione dell'utente o al valore di default
        applyTheme(selectedTheme)
        setContentView(R.layout.settings_activity)
        if (savedInstanceState == null) {
            supportFragmentManager.beginTransaction()
                .replace(R.id.settings, SettingsFragment())
                .commit()
        }
        val actionBar = supportActionBar
        actionBar?.setDisplayHomeAsUpEnabled(true)
    }

    private fun applyTheme(selectedTheme: String?) {
        val desiredMode = getNightMode(selectedTheme)
        if (desiredMode != AppCompatDelegate.getDefaultNightMode()) {
            AppCompatDelegate.setDefaultNightMode(desiredMode)
        }
    }

    private fun getNightMode(selectedTheme: String?): Int {
        return when (selectedTheme!!.lowercase(Locale.getDefault())) {
            "dark" -> AppCompatDelegate.MODE_NIGHT_YES
            "light" -> AppCompatDelegate.MODE_NIGHT_NO
            "system" -> AppCompatDelegate.MODE_NIGHT_FOLLOW_SYSTEM
            else -> AppCompatDelegate.MODE_NIGHT_FOLLOW_SYSTEM
        }
    }
}


