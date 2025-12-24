import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:new_app/main.dart';

void main() {
  testWidgets('Home screen has correct buttons', (WidgetTester tester) async {
    // Build app
    await tester.pumpWidget(const MyApp());

    // Verify title
    expect(find.text('Bilancio'), findsOneWidget);

    // Verify buttons
    expect(find.text('CONTO ECONOMICO'), findsOneWidget);
    expect(find.text('STATO PATRIMONIALE'), findsOneWidget);
    expect(find.text('FILE SALVATI'), findsOneWidget);
  });

  testWidgets('Navigation to Conto Economico and input fields work', (WidgetTester tester) async {
    await tester.pumpWidget(const MyApp());

    // Tap Conto Economico
    await tester.tap(find.text('CONTO ECONOMICO'));
    await tester.pumpAndSettle();

    // Verify Screen Title
    expect(find.text('Conto Economico'), findsOneWidget);

    // Verify sections
    expect(find.text('A) VALORE DELLA PRODUZIONE'), findsOneWidget);

    // Enter text in a field
    // Note: We need to find the specific TextField. Our InputRow has a label.
    // We can find the TextField by looking for the one usually adjacent to the label
    // or by type. Since we have many, let's just find the first one.
    
    // We scroll to make sure items are visible if needed (ListView), but for the first item it should be visible.
    final firstInput = find.byType(TextFormField).first;
    await tester.enterText(firstInput, '1234.56');
    
    expect(find.text('1234.56'), findsOneWidget);
  });

  testWidgets('Navigation to Stato Patrimoniale works', (WidgetTester tester) async {
    await tester.pumpWidget(const MyApp());

    // Tap Stato Patrimoniale
    await tester.tap(find.text('STATO PATRIMONIALE'));
    await tester.pumpAndSettle();

    // Verify Screen Title
    expect(find.text('Stato Patrimoniale'), findsOneWidget);
    
    // Verify first header
    expect(find.text('ATTIVO'), findsOneWidget);
  });
}
