import 'package:flutter_test/flutter_test.dart';
import 'package:new_app/models/financial_data.dart';
import 'package:new_app/services/excel_service.dart';

void main() {
  group('ExcelService Tests', () {
    final ExcelService service = ExcelService();

    test('generateContoEconomico returns valid bytes for empty data', () {
      final data = ContoEconomicoData();
      final bytes = service.generateContoEconomico(data);
      
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, true);
    });

    test('generateContoEconomico returns valid bytes for populated data', () {
      final data = ContoEconomicoData(
        ricaviVendite: 1000.0,
        costiMateriePrime: 500.0,
        ammortamentoImmobMateriali: 100.0,
      );
      final bytes = service.generateContoEconomico(data);
      
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, true);
    });

    test('generateStatoPatrimoniale returns valid bytes for empty data', () {
      final data = StatoPatrimonialeData();
      final bytes = service.generateStatoPatrimoniale(data);
      
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, true);
    });

    test('generateStatoPatrimoniale returns valid bytes for populated data', () {
      final data = StatoPatrimonialeData(
        creditiVersoSoci: 10000.0,
        terreniFabbricati: 50000.0,
        capitaleSociale: 20000.0,
      );
      final bytes = service.generateStatoPatrimoniale(data);
      
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, true);
    });
  });
}
