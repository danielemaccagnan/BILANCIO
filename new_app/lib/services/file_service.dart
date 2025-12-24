import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';

class FileService {
  Future<String?> saveExcelFile(String fileName, List<int> bytes) async {
    try {
      // Request storage permissions
      if (Platform.isAndroid) {
        var status = await Permission.storage.status;
        if (!status.isGranted) {
          status = await Permission.storage.request();
          if (!status.isGranted) {
            // For Android 11+ (scoped storage), we might not need explicit permission 
            // if saving to app-specific directories, but for Documents we might.
             // Try manageExternalStorage if needed, but usually app-specific external is safer.
             // Fallback to manageExternal for older apps logic is risky. 
             // Let's assume standard storage permission is enough or we save to app documents.
          }
        }
      }

      // Get directory
      Directory? directory;
      if (Platform.isAndroid) {
         // Use external files dir to make it accessible 
         directory = await getExternalStorageDirectory(); 
         // Or generic documents if possible, but that's harder on newer Android.
         // Let's stick to getExternalStorageDirectory which is typically /sdcard/Android/data/com.example.../files
         // To make it visible to user in "Documents/BILANCIO" like the original app attempted:
         // The original app used context.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "BILANCIO"
         // path_provider's getExternalStorageDirectory() gives the files root.
         // We can try to navigate up or use a different path if we want public visibility, 
         // but for modern Android, easiest is getApplicationDocumentsDirectory or getExternalStorageDirectory.
      } else {
        directory = await getApplicationDocumentsDirectory();
      }

      if (directory == null) return null;

      final String path = '${directory.path}/BILANCIO';
      final Directory folder = Directory(path);
      if (!await folder.exists()) {
        await folder.create(recursive: true);
      }

      // Unique filename
      String safeName = fileName.endsWith('.xlsx') ? fileName : '$fileName.xlsx';
      File file = File('$path/$safeName');
      int counter = 1;
      String nameWithoutExt = safeName.replaceAll('.xlsx', '');
      
      while (await file.exists()) {
        file = File('$path/$nameWithoutExt($counter).xlsx');
        counter++;
      }

      await file.writeAsBytes(bytes);
      return file.path;
    } catch (e) {
      print("Error saving file: $e");
      return null;
    }
  }

  Future<List<FileSystemEntity>> getSavedFiles() async {
    try {
      Directory? directory;
       if (Platform.isAndroid) {
         directory = await getExternalStorageDirectory(); 
      } else {
        directory = await getApplicationDocumentsDirectory();
      }
      
      if (directory == null) return [];
      final String path = '${directory.path}/BILANCIO';
      final Directory folder = Directory(path);
      
      if (await folder.exists()) {
        return folder.listSync().where((entity) => entity.path.endsWith('.xlsx')).toList();
      }
      return [];
    } catch (e) {
      return [];
    }
  }
}
