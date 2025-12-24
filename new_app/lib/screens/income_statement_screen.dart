import 'package:flutter/material.dart';
import '../models/financial_data.dart';
import '../components/input_row.dart';
import '../services/excel_service.dart';
import '../services/file_service.dart';

class ContoEconomicoScreen extends StatefulWidget {
  const ContoEconomicoScreen({Key? key}) : super(key: key);

  @override
  _ContoEconomicoScreenState createState() => _ContoEconomicoScreenState();
}

class _ContoEconomicoScreenState extends State<ContoEconomicoScreen> {
  final ContoEconomicoData _data = ContoEconomicoData();
  final ExcelService _excelService = ExcelService();
  final FileService _fileService = FileService();
  bool _isSaving = false;

  void _saveFile() async {
    setState(() => _isSaving = true);
    
    // Generate bytes
    List<int>? bytes = _excelService.generateContoEconomico(_data);
    
    if (bytes != null) {
      // Save file
      String? path = await _fileService.saveExcelFile('contoeconomico.xlsx', bytes);
      
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

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Conto Economico'),
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
                  const Text("A) VALORE DELLA PRODUZIONE", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                  InputRow(
                    label: "1) Ricavi delle vendite e delle prestazioni",
                    onChanged: (val) => _data.ricaviVendite = double.tryParse(val),
                  ),
                  InputRow(
                    label: "2) Variazione rimanenze di prodotti...",
                    onChanged: (val) => _data.variazRimanenzeProdotti = double.tryParse(val),
                  ),
                  InputRow(
                    label: "3) Variazione dei lavori in corso su ordinazione",
                    onChanged: (val) => _data.variazLavoriCorso = double.tryParse(val),
                  ),
                  InputRow(
                    label: "4) Incrementi di immobilizzazioni per lavori interni",
                    onChanged: (val) => _data.incrementiImmobilizzazioni = double.tryParse(val),
                  ),
                  InputRow(
                    label: "5) Altri ricavi e proventi",
                    onChanged: (val) => _data.altriRicavi = double.tryParse(val),
                  ),
                  const Divider(),
                  const Text("B) COSTI DELLA PRODUZIONE", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                  InputRow(
                    label: "6) Costi per acquisti di materie prime...",
                    onChanged: (val) => _data.costiMateriePrime = double.tryParse(val),
                  ),
                  InputRow(
                    label: "7) Costi per servizi",
                    onChanged: (val) => _data.costiServizi = double.tryParse(val),
                  ),
                  InputRow(
                    label: "8) Costi per godimento di beni di terzi",
                    onChanged: (val) => _data.costiGodimentoBeni = double.tryParse(val),
                  ),
                  const Text("9) Costi per il personale", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(
                    label: "a) Salari e stipendi",
                    onChanged: (val) => _data.salariStipendi = double.tryParse(val),
                  ),
                  InputRow(
                    label: "b) Oneri sociali",
                    onChanged: (val) => _data.oneriSociali = double.tryParse(val),
                  ),
                  InputRow(
                    label: "c) TFR",
                    onChanged: (val) => _data.tfr = double.tryParse(val),
                  ),
                  InputRow(
                    label: "d) Trattamento di quiescenza",
                    onChanged: (val) => _data.quiescenza = double.tryParse(val),
                  ),
                  InputRow(
                    label: "e) Altri costi",
                    onChanged: (val) => _data.altriCostiPersonale = double.tryParse(val),
                  ),
                  const Text("10) Ammortamenti e svalutazioni", style: TextStyle(fontWeight: FontWeight.bold)),
                  InputRow(
                    label: "a) Amm. immob. immateriali",
                    onChanged: (val) => _data.ammortamentoImmobImmateriali = double.tryParse(val),
                  ),
                  InputRow(
                    label: "b) Amm. immob. materiali",
                    onChanged: (val) => _data.ammortamentoImmobMateriali = double.tryParse(val),
                  ),
                  InputRow(
                    label: "c) Altre svalutazioni immob.",
                    onChanged: (val) => _data.altreSvalutazioniImmob = double.tryParse(val),
                  ),
                  InputRow(
                    label: "d) Svalutazione crediti",
                    onChanged: (val) => _data.svalutazioneCrediti = double.tryParse(val),
                  ),
                  InputRow(
                    label: "11) Variazione rimanenze materie prime...",
                    onChanged: (val) => _data.variazRimanenzeMaterie = double.tryParse(val),
                  ),
                  InputRow(
                    label: "12) Accantonamenti per rischi",
                    onChanged: (val) => _data.accantonamentiRischi = double.tryParse(val),
                  ),
                  InputRow(
                    label: "13) Altri accantonamenti",
                    onChanged: (val) => _data.altriAccantonamenti = double.tryParse(val),
                  ),
                  InputRow(
                    label: "14) Oneri diversi di gestione",
                    onChanged: (val) => _data.oneriDiversi = double.tryParse(val),
                  ),
                  const Divider(),
                  const Text("C) PROVENTI E ONERI FINANZIARI", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                  Text("15) Proventi delle partecipazioni"),
                  InputRow(label: "a) imprese controllate", onChanged: (val) => _data.proventiImpreseControllate = double.tryParse(val)),
                  InputRow(label: "b) imprese collegate", onChanged: (val) => _data.proventiImpreseCollegate = double.tryParse(val)),
                  InputRow(label: "c) imprese controllanti", onChanged: (val) => _data.proventiImpreseControllanti = double.tryParse(val)),
                  InputRow(label: "d) sottoposte controllo", onChanged: (val) => _data.proventiImpreseSottoposteControllo = double.tryParse(val)),
                  InputRow(label: "e) altre imprese", onChanged: (val) => _data.proventiAltreImprese = double.tryParse(val)),
                  
                  InputRow(label: "16) Altri proventi finanziari", onChanged: (val) => _data.altriProventiFinanziari = double.tryParse(val)),
                  InputRow(label: "17) Interessi ed altri oneri", onChanged: (val) => _data.interessiOneriFinanziari = double.tryParse(val)),
                  
                  const Divider(),
                  Text("D) RETTIFICHE DI VALORE", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                  Text("18) Rivalutazioni"),
                  InputRow(label: "a) partecipazioni", onChanged: (val) => _data.rivalutazionePartecipazioni = double.tryParse(val)),
                  InputRow(label: "b) immob. finanziarie", onChanged: (val) => _data.rivalutazioneImmobFinanziarie = double.tryParse(val)),
                  InputRow(label: "c) titoli circolante", onChanged: (val) => _data.rivalutazioneTitoliAttivo = double.tryParse(val)),
                  InputRow(label: "d) derivati", onChanged: (val) => _data.rivalutazioneStrumentiDerivati = double.tryParse(val)),
                  
                  Text("19) Svalutazioni"),
                  InputRow(label: "a) partecipazioni", onChanged: (val) => _data.svalutazionePartecipazioni = double.tryParse(val)),
                  InputRow(label: "b) immob. finanziarie", onChanged: (val) => _data.svalutazioneImmobFinanziarie = double.tryParse(val)),
                  InputRow(label: "c) titoli circolante", onChanged: (val) => _data.svalutazioneTitoliAttivo = double.tryParse(val)),
                  InputRow(label: "d) derivati", onChanged: (val) => _data.svalutazioneStrumentiDerivati = double.tryParse(val)),

                  const Divider(),
                  Text("E) PROVENTI E ONERI STRAORDINARI", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                  InputRow(label: "20) Proventi straordinari", onChanged: (val) => _data.proventiStraordinari = double.tryParse(val)),
                  InputRow(label: "21) Oneri straordinari", onChanged: (val) => _data.oneriStraordinari = double.tryParse(val)),

                  const Divider(),
                  InputRow(label: "22) Imposte sul reddito", onChanged: (val) => _data.imposteReddito = double.tryParse(val), isBold: true),
                  
                  const SizedBox(height: 30),
                  ElevatedButton(
                    onPressed: _saveFile,
                    style: ElevatedButton.styleFrom(padding: EdgeInsets.all(16)),
                    child: const Text("SALVA CONTO ECONOMICO"),
                  ),
                  const SizedBox(height: 50),
                ],
              ),
            ),
    );
  }
}
