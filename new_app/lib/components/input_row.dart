import 'package:flutter/material.dart';

class InputRow extends StatelessWidget {
  final String label;
  final ValueChanged<String> onChanged;
  final String? initialValue;
  final bool isBold;

  const InputRow({
    Key? key,
    required this.label,
    required this.onChanged,
    this.initialValue,
    this.isBold = false,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4.0),
      child: Row(
        children: [
          Expanded(
            flex: 3,
            child: Text(
              label,
              style: TextStyle(
                fontWeight: isBold ? FontWeight.bold : FontWeight.normal,
              ),
            ),
          ),
          SizedBox(width: 8),
          Expanded(
            flex: 1,
            child: TextFormField(
              initialValue: initialValue,
              keyboardType: TextInputType.numberWithOptions(decimal: true),
              decoration: InputDecoration(
                border: OutlineInputBorder(),
                contentPadding: EdgeInsets.symmetric(horizontal: 8, vertical: 0),
              ),
              onChanged: onChanged,
            ),
          ),
        ],
      ),
    );
  }
}
