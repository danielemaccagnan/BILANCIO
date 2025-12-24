class ContoEconomicoData {
  // A) Valore della produzione
  double? ricaviVendite; // 1)
  double? variazRimanenzeProdotti; // 2)
  double? variazLavoriCorso; // 3)
  double? incrementiImmobilizzazioni; // 4)
  double? altriRicavi; // 5)

  // B) Costi della produzione
  double? costiMateriePrime; // 6)
  double? costiServizi; // 7)
  double? costiGodimentoBeni; // 8)
  
  // 9) Personale
  double? salariStipendi; // a)
  double? oneriSociali; // b)
  double? tfr; // c)
  double? quiescenza; // d)
  double? altriCostiPersonale; // e)

  // 10) Ammortamenti e svalutazioni
  double? ammortamentoImmobImmateriali; // a)
  double? ammortamentoImmobMateriali; // b)
  double? altreSvalutazioniImmob; // c)
  double? svalutazioneCrediti; // d)

  double? variazRimanenzeMaterie; // 11)
  double? accantonamentiRischi; // 12)
  double? altriAccantonamenti; // 13)
  double? oneriDiversi; // 14)

  // C) Proventi e oneri finanziari
  // 15) Proventi partecipazioni
  double? proventiImpreseControllate; // a)
  double? proventiImpreseCollegate; // b)
  double? proventiImpreseControllanti; // c)
  double? proventiImpreseSottoposteControllo; // d)
  double? proventiAltreImprese; // e)

  double? altriProventiFinanziari; // 16)
  double? interessiOneriFinanziari; // 17)

  // D) Rettifiche valore attività finanziarie
  // 18) Rivalutazioni
  double? rivalutazionePartecipazioni; // a)
  double? rivalutazioneImmobFinanziarie; // b)
  double? rivalutazioneTitoliAttivo; // c)
  double? rivalutazioneStrumentiDerivati; // d)

  // 19) Svalutazioni
  double? svalutazionePartecipazioni; // a)
  double? svalutazioneImmobFinanziarie; // b)
  double? svalutazioneTitoliAttivo; // c)
  double? svalutazioneStrumentiDerivati; // d)

  // E) Proventi e oneri straordinari (Legacy in 2024? Keeping as per original app)
  double? proventiStraordinari; // 20)
  double? oneriStraordinari; // 21)

  double? imposteReddito; // 22)

  ContoEconomicoData({
    this.ricaviVendite,
    this.variazRimanenzeProdotti,
    this.variazLavoriCorso,
    this.incrementiImmobilizzazioni,
    this.altriRicavi,
    this.costiMateriePrime,
    this.costiServizi,
    this.costiGodimentoBeni,
    this.salariStipendi,
    this.oneriSociali,
    this.tfr,
    this.quiescenza,
    this.altriCostiPersonale,
    this.ammortamentoImmobImmateriali,
    this.ammortamentoImmobMateriali,
    this.altreSvalutazioniImmob,
    this.svalutazioneCrediti,
    this.variazRimanenzeMaterie,
    this.accantonamentiRischi,
    this.altriAccantonamenti,
    this.oneriDiversi,
    this.proventiImpreseControllate,
    this.proventiImpreseCollegate,
    this.proventiImpreseControllanti,
    this.proventiImpreseSottoposteControllo,
    this.proventiAltreImprese,
    this.altriProventiFinanziari,
    this.interessiOneriFinanziari,
    this.rivalutazionePartecipazioni,
    this.rivalutazioneImmobFinanziarie,
    this.rivalutazioneTitoliAttivo,
    this.rivalutazioneStrumentiDerivati,
    this.svalutazionePartecipazioni,
    this.svalutazioneImmobFinanziarie,
    this.svalutazioneTitoliAttivo,
    this.svalutazioneStrumentiDerivati,
    this.proventiStraordinari,
    this.oneriStraordinari,
    this.imposteReddito,
  });
}

class StatoPatrimonialeData {
  // ATTIVO
  // A) Crediti verso soci
  double? creditiVersoSoci; // A)

  // B) Immobilizzazioni
  // I) Immateriali
  double? costiImpianto; // 1)
  double? costiSviluppo; // 2)
  double? dirittiBrevetto; // 3)
  double? concessioniLicenze; // 4)
  double? avviamento; // 5)
  double? immobilizzazioniCorso; // 6)
  double? altreImmobilizzazioniImmateriali; // 7)

  // II) Materiali
  double? terreniFabbricati; // 1)
  double? impiantiMacchinari; // 2)
  double? attrezzatureIndustriali; // 3)
  double? altriBeniMateriali; // 4)
  double? immobilizzazioniMaterialiCorso; // 5)

  // III) Finanziarie
  double? partecipazioniImpreseControllate; // 1a
  double? partecipazioniImpreseCollegate; // 1b
  double? partecipazioniImpreseControllanti; // 1c
  double? partecipazioniAltreImprese; // 1d
  double? creditiImpreseControllate; // 2a
  double? creditiImpreseCollegate; // 2b
  double? creditiImpreseControllanti; // 2c
  double? creditiAltreImprese; // 2d
  double? altriTitoli; // 3
  double? azioniProprie; // 4

  // C) Attivo Circolante
  // I) Rimanenze
  double? rimanenzeMateriePrime; // 1)
  double? rimanenzeProdottiCorso; // 2)
  double? rimanenzeLavoriCorso; // 3)
  double? rimanenzeProdottiFiniti; // 4)
  double? accontiRimanenze; // 5)

  // II) Crediti
  double? creditiClienti; // 1)
  double? creditiClientiImpreseControllate; // 2)
  double? creditiClientiImpreseCollegate; // 3)
  double? creditiClientiImpreseControllanti; // 4)
  double? creditiTributari; // 5-bis)
  double? imposteAnticipate; // 5-ter)
  double? creditiVersoAltri; // 5-quater)

  // III) Attività finanziarie che non costituiscono immobilizzazioni
  double? partecipazioniNonImmobilizzate; // 1)
  double? altriTitoliNonImmobilizzati; // 6)

  // IV) Disponibilità liquide
  double? depositiBancari; // 1)
  double? assegni; // 2)
  double? denaroCassa; // 3)

  // D) Ratei e Risconti
  double? rateiRiscontiAttivi; // D)

  // PASSIVO
  // A) Patrimonio Netto
  double? capitaleSociale; // I
  double? riservaSovrapprezzo; // II
  double? riservaRivalutazione; // III
  double? riservaLegale; // IV
  double? riserveStatutarie; // V
  double? altreRiserve; // VI
  double? riservaCoperturaFlussi; // VII
  double? utiliPortatiNuovo; // VIII
  double? utileEsercizio; // IX
  double? riservaNegativaAzioniProprie; // X

  // B) Fondi per rischi e oneri
  double? fondiPensione; // 1)
  double? fondiImposte; // 2)
  double? strumentiFinanziariDerivatiPassivi; // 3)
  double? altriFondi; // 4)

  // C) TFR
  double? tfrPassivo; // C)

  // D) Debiti
  double? obbligazioni; // 1)
  double? obbligazioniConvertibili; // 2)
  double? debitiSoci; // 3)
  double? debitiBanche; // 4)
  double? debitiAltriFinanziatori; // 5)
  double? accontiDebiti; // 6)
  double? debitiFornitori; // 7)
  double? titoliCredito; // 8)
  double? debitiImpreseControllate; // 9)
  double? debitiImpreseCollegate; // 10)
  double? debitiImpreseControllanti; // 11)
  double? debitiTributari; // 12)
  double? debitiPrevidenziali; // 13)
  double? altriDebiti; // 14)

  // E) Ratei e Risconti
  double? rateiRiscontiPassivi; // E)

  StatoPatrimonialeData({
    this.creditiVersoSoci,
    this.costiImpianto,
    this.costiSviluppo,
    this.dirittiBrevetto,
    this.concessioniLicenze,
    this.avviamento,
    this.immobilizzazioniCorso,
    this.altreImmobilizzazioniImmateriali,
    this.terreniFabbricati,
    this.impiantiMacchinari,
    this.attrezzatureIndustriali,
    this.altriBeniMateriali,
    this.immobilizzazioniMaterialiCorso,
    this.partecipazioniImpreseControllate,
    this.partecipazioniImpreseCollegate,
    this.partecipazioniImpreseControllanti,
    this.partecipazioniAltreImprese,
    this.creditiImpreseControllate,
    this.creditiImpreseCollegate,
    this.creditiImpreseControllanti,
    this.creditiAltreImprese,
    this.altriTitoli,
    this.azioniProprie,
    this.rimanenzeMateriePrime,
    this.rimanenzeProdottiCorso,
    this.rimanenzeLavoriCorso,
    this.rimanenzeProdottiFiniti,
    this.accontiRimanenze,
    this.creditiClienti,
    this.creditiClientiImpreseControllate,
    this.creditiClientiImpreseCollegate,
    this.creditiClientiImpreseControllanti,
    this.creditiTributari,
    this.imposteAnticipate,
    this.creditiVersoAltri,
    this.partecipazioniNonImmobilizzate,
    this.altriTitoliNonImmobilizzati,
    this.depositiBancari,
    this.assegni,
    this.denaroCassa,
    this.rateiRiscontiAttivi,
    this.capitaleSociale,
    this.riservaSovrapprezzo,
    this.riservaRivalutazione,
    this.riservaLegale,
    this.riserveStatutarie,
    this.altreRiserve,
    this.riservaCoperturaFlussi,
    this.utiliPortatiNuovo,
    this.utileEsercizio,
    this.riservaNegativaAzioniProprie,
    this.fondiPensione,
    this.fondiImposte,
    this.strumentiFinanziariDerivatiPassivi,
    this.altriFondi,
    this.tfrPassivo,
    this.obbligazioni,
    this.obbligazioniConvertibili,
    this.debitiSoci,
    this.debitiBanche,
    this.debitiAltriFinanziatori,
    this.accontiDebiti,
    this.debitiFornitori,
    this.titoliCredito,
    this.debitiImpreseControllate,
    this.debitiImpreseCollegate,
    this.debitiImpreseControllanti,
    this.debitiTributari,
    this.debitiPrevidenziali,
    this.altriDebiti,
    this.rateiRiscontiPassivi,
  });
}
