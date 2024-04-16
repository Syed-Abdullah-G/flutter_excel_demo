import 'dart:io';
import 'package:external_path/external_path.dart';
import 'package:dio/dio.dart';
import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:downloads_path_provider_28/downloads_path_provider_28.dart';

//import 'package:dio/dio.dart';
void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
        title: 'Flutter Demo',
        theme: ThemeData(
          colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
          useMaterial3: true,
        ),
        home: const MyHomePage());
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String _externalStorageDirectory = '';

  @override
  void initState() {
    // TODO: implement initState
    super.initState();
    _getExternalStorageDirectory();
  }

  Future<void> _getExternalStorageDirectory() async {
    try {
      String directory = await ExternalPath.getExternalStoragePublicDirectory(
        ExternalPath.DIRECTORY_DOWNLOADS,
      );
      setState(() {
        _externalStorageDirectory = directory;
      });
    } catch (e) {
      print('Error getting external storage directory: $e');
    }
  }

  final  _storagePermission = Permission.storage;

  Future<void> createExcel() async {
    final status = await _storagePermission.request();
    if (status.isGranted) {
      print('Storage permission given...');
      final Workbook workbook = new Workbook();
      final Worksheet sheet = workbook.worksheets[0];
      sheet.getRangeByName('A1').setText('Syed Abdullah');
      sheet.getRangeByName('A2').setNumber(18);
      final List<int> bytes = workbook.saveAsStream();
      workbook.dispose();

      if (_externalStorageDirectory.isNotEmpty) {
        var time = DateTime.now().millisecondsSinceEpoch;
        var path = '$_externalStorageDirectory/image-$time.xlsx';
        var file = File(path);
        await file.writeAsBytes(bytes);
        print('Excel file saved at: $path');
      } else {
        print('External storage directory is not available');
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: ElevatedButton(
            onPressed: createExcel, child: Text('Save Excel File')),
      ),
    );
  }
}
