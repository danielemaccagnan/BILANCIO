package com.example.bilancio

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import androidx.preference.ListPreference
import androidx.preference.Preference
import androidx.preference.PreferenceFragmentCompat
import androidx.appcompat.app.AppCompatDelegate

class SettingsFragment : PreferenceFragmentCompat() {

    override fun onCreatePreferences(savedInstanceState: Bundle?, rootKey: String?) {
        setPreferencesFromResource(R.xml.root_preferences, rootKey)

        // Trova la preferenza del tema utilizzando la sua chiave
        val themePreference: ListPreference? = findPreference("theme")
        themePreference?.setOnPreferenceChangeListener { _, newValue ->
            val themeOption = newValue as String
            updateTheme(themeOption)
            true
        }

        // Trova e gestisci la preferenza "Condividi"
        val sharePreference: Preference? = findPreference("share")
        sharePreference?.setOnPreferenceClickListener {
            shareApp()
            true
        }

        // Trova e gestisci la preferenza "Recensisci questa app"
        val reviewPreference: Preference? = findPreference("review")
        reviewPreference?.setOnPreferenceClickListener {
            reviewApp()
            true
        }
    }

    private fun updateTheme(themeOption: String) {
        when (themeOption) {
            "light" -> AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_NO)
            "dark" -> AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_YES)
            else -> AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_FOLLOW_SYSTEM)
        }
    }

    private fun shareApp() {
        val shareIntent = Intent(Intent.ACTION_SEND).apply {
            type = "text/plain"
            val shareBody = "Dai un'occhiata a Bilancio, un'app fantastica per aiutarti a risolvere gli esercizi di economia!"
            // Aggiungi qui il link diretto al Google Play Store se disponibile
            val shareSub = "Prova Bilancio!"
            putExtra(Intent.EXTRA_SUBJECT, shareSub)
            putExtra(Intent.EXTRA_TEXT, shareBody)
        }
        startActivity(Intent.createChooser(shareIntent, "Condividi con"))
    }

    private fun reviewApp() {
        // Sostituisci "com.example.bilancio" con l'ID effettivo del pacchetto della tua app
        val uri = Uri.parse("market://details?id=com.example.bilancio")
        val goToMarket = Intent(Intent.ACTION_VIEW, uri).apply {
            addFlags(Intent.FLAG_ACTIVITY_NO_HISTORY or
                    Intent.FLAG_ACTIVITY_NEW_DOCUMENT or
                    Intent.FLAG_ACTIVITY_MULTIPLE_TASK)
        }
        try {
            startActivity(goToMarket)
        } catch (e: Exception) {
            startActivity(Intent(Intent.ACTION_VIEW,
                Uri.parse("http://play.google.com/store/apps/details?id=com.example.bilancio")))
        }
    }
}


