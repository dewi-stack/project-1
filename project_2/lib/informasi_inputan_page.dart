// Pastikan import tidak berubah
import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:intl/intl.dart';
import 'package:syncfusion_flutter_datagrid/datagrid.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:syncfusion_flutter_core/theme.dart';
import 'package:universal_html/html.dart' as html;
import 'package:shared_preferences/shared_preferences.dart';

import 'login_page.dart';

class InformasiInputanPage extends StatefulWidget {
  final String token;

  const InformasiInputanPage({super.key, required this.token});

  @override
  State<InformasiInputanPage> createState() => _InformasiInputanPageState();
}

class _InformasiInputanPageState extends State<InformasiInputanPage> {
  List<Inputan> _inputanList = [];
  List<Inputan> _filteredList = [];
  bool _isLoading = true;
  late InputanDataSource _dataSource;

  final int _rowsPerPage = 7;
  DateTimeRange? _selectedDateRange;

  @override
  void initState() {
    super.initState();
    fetchData();
  }

  Future<void> fetchData() async {
    final url = Uri.parse('http://apppenjualan791.my.id/api/admin/karyawans');

    try {
      final response = await http.get(url, headers: {'Authorization': 'Bearer ${widget.token}'});

      if (response.statusCode == 200) {
        final jsonData = json.decode(response.body);

        if (jsonData is List) {
          _inputanList = jsonData.map<Inputan>((item) {
            final hargaJual = double.tryParse(item['harga_jual'].toString()) ?? 0.0;
            final modal = double.tryParse(item['modal'].toString()) ?? 0.0;
            final profit = hargaJual - modal;
            final createdAt = DateTime.tryParse(item['created_at'] ?? '') ?? DateTime.now();

            return Inputan(
              nama: item['nama'] ?? '-',
              hargaJual: hargaJual,
              modal: modal,
              totalProfit: profit,
              tanggal: createdAt,
            );
          }).toList();

          _inputanList.sort((a, b) => b.tanggal.compareTo(a.tanggal));
        }

        setState(() {
          _filteredList = _inputanList;
          _dataSource = InputanDataSource(inputanList: _filteredList);
          _isLoading = false;
        });
      } else {
        throw Exception('Gagal mengambil data: ${response.statusCode}');
      }
    } catch (e) {
      print('Error: $e');
      setState(() => _isLoading = false);
    }
  }

 void filterByDateRange(DateTimeRange? range) {
    if (range == null) {
      setState(() {
        _filteredList = _inputanList;
        _dataSource = InputanDataSource(inputanList: _filteredList);
      });
      return;
    }

    final start = DateTime(range.start.year, range.start.month, range.start.day);
    final end = DateTime(range.end.year, range.end.month, range.end.day, 23, 59, 59);

    setState(() {
      _filteredList = _inputanList.where((inputan) {
        return inputan.tanggal.isAtSameMomentAs(start) ||
              inputan.tanggal.isAtSameMomentAs(end) ||
              (inputan.tanggal.isAfter(start) && inputan.tanggal.isBefore(end));
      }).toList();

      _dataSource = InputanDataSource(inputanList: _filteredList);
    });
  }

  Future<void> exportToExcel() async {
    final workbook = xlsio.Workbook();
    final sheet = workbook.worksheets[0];

    final headerStyle = workbook.styles.add('headerStyle');
    headerStyle.bold = true;
    headerStyle.backColor = '#DCE6F1';
    headerStyle.hAlign = xlsio.HAlignType.center;
    headerStyle.borders.all.lineStyle = xlsio.LineStyle.thin;

    final currencyStyle = workbook.styles.add('CurrencyStyle');
    currencyStyle.numberFormat = r'_([$Rp-421]* #,##0_)';

    final headers = ['Nama', 'Harga Jual', 'Modal', 'Untung', 'Tanggal'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(1, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle.bold = true;
    }

    double totalProfit = 0;

    for (int i = 0; i < _filteredList.length; i++) {
      final data = _filteredList[i];
      final row = i + 2;

      sheet.getRangeByIndex(row, 1).setText(data.nama);
      sheet.getRangeByIndex(row, 2)..setNumber(data.hargaJual)..cellStyle = currencyStyle;
      sheet.getRangeByIndex(row, 3)..setNumber(data.modal)..cellStyle = currencyStyle;
      sheet.getRangeByIndex(row, 4)..setNumber(data.totalProfit)..cellStyle = currencyStyle;
      sheet.getRangeByIndex(row, 5).setText(InputanDataSource.formatTanggal(data.tanggal));

      totalProfit += data.totalProfit;
    }

    final lastRow = _filteredList.length + 2;
    sheet.getRangeByIndex(lastRow, 3)
      ..setText('Total Profit')
      ..cellStyle.bold = true;

    sheet.getRangeByIndex(lastRow, 4)
      ..setNumber(totalProfit)
      ..cellStyle = currencyStyle
      ..cellStyle.bold = true;

    final bytes = workbook.saveAsStream();
    workbook.dispose();

    final String namaFile = 'laporan_informasi_inputan_karyawan.xlsx';

    if (kIsWeb) {
      final blob = html.Blob([Uint8List.fromList(bytes)]);
      final url = html.Url.createObjectUrlFromBlob(blob);
      final anchor = html.AnchorElement(href: url)
        ..setAttribute("download", namaFile)
        ..click();
      html.Url.revokeObjectUrl(url);
    } else {
      bool permissionGranted = false;

      if (Platform.isAndroid) {
        var status = await Permission.manageExternalStorage.status;
        if (!status.isGranted) {
          status = await Permission.manageExternalStorage.request();
        }
        permissionGranted = status.isGranted;
      } else {
        var status = await Permission.storage.status;
        if (!status.isGranted) {
          status = await Permission.storage.request();
        }
        permissionGranted = status.isGranted;
      }

      if (!permissionGranted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Izin penyimpanan tidak diberikan')),
        );
        return;
      }

      try {
        final dir = await getExternalStorageDirectory();
        String newPath = "";
        List<String> folders = dir!.path.split("/");
        for (int i = 1; i < folders.length; i++) {
          String folder = folders[i];
          if (folder == "Android") break;
          newPath += "/$folder";
        }
        newPath += "/Download";
        final path = '$newPath/$namaFile';

        final file = File(path);
        await file.writeAsBytes(bytes, flush: true);

        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Berhasil diekspor ke: $path')),
        );
      } catch (e) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Gagal menyimpan file: $e')),
        );
      }
    }
  }

  Future<void> logout() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.remove('token');
    await prefs.remove('role');

    if (!mounted) return;
    Navigator.pushAndRemoveUntil(
      context,
      MaterialPageRoute(builder: (context) => const LoginPage()),
      (route) => false,
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Informasi Inputan Data'),
        centerTitle: true,
        backgroundColor: Colors.indigo,
        foregroundColor: Colors.white,
        actions: [
          IconButton(
            icon: const Icon(Icons.logout),
            tooltip: 'Logout',
            onPressed: logout,
          ),
        ],
      ),
      body: _isLoading
          ? const Center(child: CircularProgressIndicator())
          : Padding(
              padding: const EdgeInsets.symmetric(horizontal: 12.0),
              child: Card(
                elevation: 4,
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(12),
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Padding(
                      padding: EdgeInsets.all(16.0),
                      child: Text(
                        'Data Inputan Karyawan',
                        style: TextStyle(
                          fontSize: 18,
                          fontWeight: FontWeight.bold,
                          color: Colors.black,
                        ),
                      ),
                    ),
                    Padding(
                      padding: const EdgeInsets.symmetric(horizontal: 16.0),
                      child: Wrap(
                        spacing: 8,
                        runSpacing: 8,
                        children: [
                          ElevatedButton.icon(
                            icon: const Icon(Icons.date_range),
                            label: Text(
                              _selectedDateRange == null
                                  ? 'Pilih Rentang Tanggal'
                                  : '${DateFormat('dd-MM-yy').format(_selectedDateRange!.start)} - ${DateFormat('dd-MM-yy').format(_selectedDateRange!.end)}',
                            ),
                            onPressed: () async {
                              final picked = await showDateRangePicker(
                                context: context,
                                firstDate: DateTime(DateTime.now().year - 5),
                                lastDate: DateTime(DateTime.now().year + 5),
                                initialDateRange: _selectedDateRange,
                                builder: (context, child) {
                                  return Theme(
                                    data: Theme.of(context).copyWith(
                                      colorScheme: const ColorScheme.light(
                                        primary: Colors.deepPurple,
                                        onPrimary: Colors.white,
                                        surface: Colors.white,
                                      ),
                                      textButtonTheme: TextButtonThemeData(
                                        style: TextButton.styleFrom(
                                          foregroundColor: Colors.deepPurple,
                                        ),
                                      ),
                                    ),
                                    child: child!,
                                  );
                                },
                              );

                              if (picked != null) {
                                setState(() {
                                  _selectedDateRange = picked;
                                });
                                filterByDateRange(picked);
                              }
                            },
                          ),
                          ElevatedButton.icon(
                            icon: const Icon(Icons.refresh),
                            label: const Text('Reset'),
                            onPressed: () {
                              setState(() {
                                _selectedDateRange = null;
                                filterByDateRange(null);
                              });
                            },
                            style: ElevatedButton.styleFrom(backgroundColor: Colors.grey),
                          ),
                          ElevatedButton.icon(
                            icon: const Icon(Icons.download),
                            label: const Text('Export ke Excel'),
                            onPressed: exportToExcel,
                            style: ElevatedButton.styleFrom(
                              backgroundColor: Colors.green,
                              foregroundColor: Colors.white,
                            ),
                          ),
                        ],
                      ),
                    ),
                    const Divider(height: 10, thickness: 1),
                    Expanded(
                      child: SfDataGridTheme(
                        data: SfDataGridThemeData(
                          headerColor: Colors.indigo.shade600,
                          gridLineColor: Colors.grey.shade300,
                        ),
                        child: SfDataGrid(
                          source: _dataSource,
                          allowSorting: true,
                          columnWidthMode: ColumnWidthMode.fill,
                          columns: [
                            GridColumn(columnName: 'nama', label: _buildHeader('Nama', Colors.white)),
                            GridColumn(columnName: 'hargaJual', label: _buildHeader('Harga Jual', Colors.white)),
                            GridColumn(columnName: 'modal', label: _buildHeader('Modal', Colors.white)),
                            GridColumn(columnName: 'totalProfit', label: _buildHeader('Untung', Colors.white)),
                            GridColumn(columnName: 'tanggal', label: _buildHeader('Tanggal', Colors.white)),
                          ],
                        ),
                      ),
                    ),
                    Container(
                      padding: const EdgeInsets.only(bottom: 16),
                      alignment: Alignment.center,
                      child: SfDataPager(
                        delegate: _dataSource,
                        pageCount: (_filteredList.length / _rowsPerPage).ceilToDouble(),
                        direction: Axis.horizontal,
                        visibleItemsCount: 5,
                      ),
                    ),
                  ],
                ),
              ),
            ),
    );
  }

  Widget _buildHeader(String text, [Color textColor = Colors.black]) {
    return Center(
      child: Text(
        text,
        style: TextStyle(
          fontWeight: FontWeight.bold,
          color: textColor,
        ),
      ),
    );
  }
}

class Inputan {
  final String nama;
  final double hargaJual;
  final double modal;
  final double totalProfit;
  final DateTime tanggal;

  Inputan({
    required this.nama,
    required this.hargaJual,
    required this.modal,
    required this.totalProfit,
    required this.tanggal,
  });
}

class InputanDataSource extends DataGridSource {
  List<DataGridRow> _inputanRows = [];

  InputanDataSource({required List<Inputan> inputanList}) {
    _inputanRows = inputanList.map<DataGridRow>((inputan) {
      return DataGridRow(cells: [
        DataGridCell<String>(columnName: 'nama', value: inputan.nama),
        DataGridCell<String>(columnName: 'hargaJual', value: formatRupiah(inputan.hargaJual)),
        DataGridCell<String>(columnName: 'modal', value: formatRupiah(inputan.modal)),
        DataGridCell<String>(columnName: 'totalProfit', value: formatRupiah(inputan.totalProfit)),
        DataGridCell<DateTime>(columnName: 'tanggal', value: inputan.tanggal),
      ]);
    }).toList();
  }

  @override
  List<DataGridRow> get rows => _inputanRows;

  @override
  DataGridRowAdapter buildRow(DataGridRow row) {
    return DataGridRowAdapter(
      cells: row.getCells().map((cell) {
        return Container(
          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
          alignment: Alignment.centerLeft,
          child: Text(
            cell.columnName == 'tanggal'
                ? formatTanggal(cell.value)
                : cell.value.toString(),
          ),
        );
      }).toList(),
    );
  }

  static String formatRupiah(double number) {
    final formatter = NumberFormat.currency(locale: 'id_ID', symbol: 'Rp ', decimalDigits: 0);
    return formatter.format(number);
  }

  static String formatTanggal(DateTime date) {
    return DateFormat('dd-MM-yy').format(date); // <- DIUBAH SESUAI PERMINTAAN
  }
}
