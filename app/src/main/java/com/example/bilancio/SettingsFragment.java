package com.example.bilancio;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import androidx.preference.ListPreference;
import androidx.preference.Preference;
import androidx.preference.PreferenceFragmentCompat;
import androidx.appcompat.app.AppCompatDelegate;

public class SettingsFragment extends PreferenceFragmentCompat {

    @Override
    public void onCreatePreferences(Bundle savedInstanceState, String rootKey) {
        setPreferencesFromResource(R.xml.root_preferences, rootKey);

        // Trova la preferenza del tema utilizzando la sua chiave
        ListPreference themePreference = findPreference("theme");
        if (themePreference != null) {
            themePreference.setOnPreferenceChangeListener(new Preference.OnPreferenceChangeListener() {
                @Override
                public boolean onPreferenceChange(Preference preference, Object newValue) {
                    String themeOption = (String) newValue;
                    updateTheme(themeOption);
                    return true;
                }
            });
        }

        // Trova e gestisci la preferenza "Condividi"
        Preference sharePreference = findPreference("share");
        if (sharePreference != null) {
            sharePreference.setOnPreferenceClickListener(preference -> {
                shareApp();
                return true;
            });
        }

        // Trova e gestisci la preferenza "Recensisci questa app"
        Preference reviewPreference = findPreference("review");
        if (reviewPreference != null) {
            reviewPreference.setOnPreferenceClickListener(preference -> {
                reviewApp();
                return true;
            });
        }
    }

    private void updateTheme(String themeOption) {
        switch (themeOption) {
            case "light":
                AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_NO);
                break;
            case "dark":
                AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_YES);
                break;
            default:
                AppCompatDelegate.setDefaultNightMode(AppCompatDelegate.MODE_NIGHT_FOLLOW_SYSTEM);
                break;
        }
    }

    private void shareApp() {
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("text/plain");
        String shareBody = "Dai un'occhiata a Bilancio, un'app fantastica per aiutarti a risolvere gli esercizi di economia!";
        // Aggiungi qui il link diretto al Google Play Store se disponibile
        String shareSub = "Prova Bilancio!";
        shareIntent.putExtra(Intent.EXTRA_SUBJECT, shareSub);
        shareIntent.putExtra(Intent.EXTRA_TEXT, shareBody);
        startActivity(Intent.createChooser(shareIntent, "Condividi con"));
    }

    private void reviewApp() {
        // Sostituisci "com.example.bilancio" con l'ID effettivo del pacchetto della tua app
        Uri uri = Uri.parse("market://details?id=com.example.bilancio");
        Intent goToMarket = new Intent(Intent.ACTION_VIEW, uri);
        // Per provare ad aprire il Google Play Store prima e, se non disponibile, aprire il browser
        goToMarket.addFlags(Intent.FLAG_ACTIVITY_NO_HISTORY |
                Intent.FLAG_ACTIVITY_NEW_DOCUMENT |
                Intent.FLAG_ACTIVITY_MULTIPLE_TASK);
        try {
            startActivity(goToMarket);
        } catch (Exception e) {
            startActivity(new Intent(Intent.ACTION_VIEW,
                    Uri.parse("http://play.google.com/store/apps/details?id=com.example.bilancio")));
        }
    }
}
