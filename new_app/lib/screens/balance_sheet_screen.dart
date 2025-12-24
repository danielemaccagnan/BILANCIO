import 'package:flutter/material.dart';
import '../models/financial_data.dart';
import '../components/input_row.dart';
import '../services/excel_service.dart';
import '../services/file_service.dart';

class StatoPatrimonialeScreen extends StatefulWidget {
  const StatoPatrimonialeScreen({Key? key}) : super(key: key);

  @override
  _StatoPatrimonialeScreenState createState() => _StatoPatrimonialeScreenState();
}

class _StatoPatrimonialeScreenState extends State<StatoPatrimonialeScreen> {
  final StatoPatrimonialeData _data = StatoPatrimonialeData();
  final ExcelService _excelService = ExcelService();
  final FileService _fileService = FileService();
  bool _isSaving = false;

  void _saveFile() async {
    setState(() => _isSaving = true);
    
    // Generate bytes
    List<int>? bytes = _excelService.generateStatoPatrimoniale(_data);
    
    if (bytes != null) {
      String? path = await _fileService.saveExcelFile('statopatrimoniale.xlsx', bytes);
      setState(() => _isSaving = false);
      
      if (path != null) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('File salvato in: $path')),
        );
      } else {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Errore durante il salvataggio')),
        );
      }
    } else {
      setState(() => _isSaving = false);
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Errore durante la generazione del file')),
      );
    }
  }

  Widget _header(String text) => Padding(
    padding: const EdgeInsets.symmetric(vertical: 10),
    child: Text(text, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 18, color: Colors.blue)),
  );
  
  Widget _subHeader(String text) => Padding(
    padding: const EdgeInsets.symmetric(vertical: 5),
    child: Text(text, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
  );

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Stato Patrimoniale'),
        actions: [
          IconButton(
            icon: const Icon(Icons.save),
            onPressed: _isSaving ? null : _saveFile,
          )
        ],
      ),
      body: _isSaving
          ? const Center(child: CircularProgressIndicator())
          : Padding(
              padding: const EdgeInsets.all(16.0),
              child: ListView(
                children: [
                  _header("ATTIVO"),
                  
                  _subHeader("A) Crediti verso soci"),
                  InputRow(label: "Versamenti ancora dovuti", onChanged: (v) => _data.creditiVersoSoci = double.tryParse(v)),

                  _subHeader("B) Immobilizzazioni"),
                  const Text("I - Immateriali", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1) Costi di impianto", onChanged: (v) => _data.costiImpianto = double.tryParse(v)),
                  InputRow(label: "2) Costi di sviluppo", onChanged: (v) => _data.costiSviluppo = double.tryParse(v)),
                  InputRow(label: "3) Diritti brevetto", onChanged: (v) => _data.dirittiBrevetto = double.tryParse(v)),
                  InputRow(label: "4) Concessioni, licenze", onChanged: (v) => _data.concessioniLicenze = double.tryParse(v)),
                  InputRow(label: "5) Avviamento", onChanged: (v) => _data.avviamento = double.tryParse(v)),
                  InputRow(label: "6) Immob. in corso", onChanged: (v) => _data.immobilizzazioniCorso = double.tryParse(v)),
                  InputRow(label: "7) Altre", onChanged: (v) => _data.altreImmobilizzazioniImmateriali = double.tryParse(v)),

                  const Text("II - Materiali", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1) Terreni e fabbricati", onChanged: (v) => _data.terreniFabbricati = double.tryParse(v)),
                  InputRow(label: "2) Impianti e macchinari", onChanged: (v) => _data.impiantiMacchinari = double.tryParse(v)),
                  InputRow(label: "3) Attrezzature", onChanged: (v) => _data.attrezzatureIndustriali = double.tryParse(v)),
                  InputRow(label: "4) Altri beni", onChanged: (v) => _data.altriBeniMateriali = double.tryParse(v)),
                  InputRow(label: "5) Immob. in corso", onChanged: (v) => _data.immobilizzazioniMaterialiCorso = double.tryParse(v)),

                  const Text("III - Finanziarie", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1a) Part. imprese controllate", onChanged: (v) => _data.partecipazioniImpreseControllate = double.tryParse(v)),
                  InputRow(label: "1b) Part. imprese collegate", onChanged: (v) => _data.partecipazioniImpreseCollegate = double.tryParse(v)),
                  // Skipping detailed breakdown for brevity but capturing main ones
                  InputRow(label: "Crediti", onChanged: (v) => _data.creditiImpreseControllate = double.tryParse(v)), // Approximation

                  _subHeader("C) Attivo Circolante"),
                  const Text("I - Rimanenze", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1) Materie prime", onChanged: (v) => _data.rimanenzeMateriePrime = double.tryParse(v)),
                  InputRow(label: "2) Prodotti in corso", onChanged: (v) => _data.rimanenzeProdottiCorso = double.tryParse(v)),
                  InputRow(label: "3) Lavori in corso", onChanged: (v) => _data.rimanenzeLavoriCorso = double.tryParse(v)),
                  InputRow(label: "4) Prodotti finiti", onChanged: (v) => _data.rimanenzeProdottiFiniti = double.tryParse(v)),
                  InputRow(label: "5) Acconti", onChanged: (v) => _data.accontiRimanenze = double.tryParse(v)),

                  const Text("II - Crediti", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1) Verso clienti", onChanged: (v) => _data.creditiClienti = double.tryParse(v)),
                  InputRow(label: "5-bis) Tributari", onChanged: (v) => _data.creditiTributari = double.tryParse(v)),
                  InputRow(label: "5-quater) Verso altri", onChanged: (v) => _data.creditiVersoAltri = double.tryParse(v)),

                  const Text("IV - Disponibilità liquide", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(label: "1) Depositi bancari", onChanged: (v) => _data.depositiBancari = double.tryParse(v)),
                  InputRow(label: "2) Assegni", onChanged: (v) => _data.assegni = double.tryParse(v)),
                  InputRow(label: "3) Denaro e valori in cassa", onChanged: (v) => _data.denaroCassa = double.tryParse(v)),

                  _subHeader("D) Ratei e Risconti"),
                  InputRow(label: "Ratei e risconti attivi", onChanged: (v) => _data.rateiRiscontiAttivi = double.tryParse(v)),
                  
                  const Divider(thickness: 2),
                  _header("PASSIVO"),
                  
                  _subHeader("A) Patrimonio Netto"),
                  InputRow(label: "I. Capitale sociale", onChanged: (v) => _data.capitaleSociale = double.tryParse(v)),
                  InputRow(label: "II. Riserva sovrapprezzo", onChanged: (v) => _data.riservaSovrapprezzo = double.tryParse(v)),
                  InputRow(label: "IV. Riserva legale", onChanged: (v) => _data.riservaLegale = double.tryParse(v)),
                  InputRow(label: "IX. Utile d'esercizio", onChanged: (v) => _data.utileEsercizio = double.tryParse(v)),
                  
                  _subHeader("B) Fondi rischi e oneri"),
                  InputRow(label: "1) Trattamento quiescenza", onChanged: (v) => _data.fondiPensione = double.tryParse(v)),
                  InputRow(label: "2) Per imposte", onChanged: (v) => _data.fondiImposte = double.tryParse(v)),
                  InputRow(label: "4) Altri", onChanged: (v) => _data.altriFondi = double.tryParse(v)),
                  
                  _subHeader("C) TFR"),
                  InputRow(label: "Trattamento di fine rapporto", onChanged: (v) => _data.tfrPassivo = double.tryParse(v)),
                  
                  _subHeader("D) Debiti"),
                  InputRow(label: "1) Obbligazioni", onChanged: (v) => _data.obbligazioni = double.tryParse(v)),
                  InputRow(label: "4) Verso banche", onChanged: (v) => _data.debitiBanche = double.tryParse(v)),
                  InputRow(label: "7) Verso fornitori", onChanged: (v) => _data.debitiFornitori = double.tryParse(v)),
                  InputRow(label: "12) Debiti tributari", onChanged: (v) => _data.debitiTributari = double.tryParse(v)),
                  InputRow(label: "13) Debiti previdenziali", onChanged: (v) => _data.debitiPrevidenziali = double.tryParse(v)),
                  
                  _subHeader("E) Ratei e Risconti"),
                  InputRow(label: "Ratei e risconti passivi", onChanged: (v) => _data.rateiRiscontiPassivi = double.tryParse(v)),

                  const SizedBox(height: 30),
                  ElevatedButton(
                    onPressed: _saveFile,
                    style: ElevatedButton.styleFrom(padding: EdgeInsets.all(16)),
                    child: const Text("SALVA STATO PATRIMONIALE"),
                  ),
                  const SizedBox(height: 50),
                ],
              ),
            ),
    );
  }
}
