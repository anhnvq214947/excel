// Tạo một Workbook mới (tương ứng với một file excel)
Workbook workbook = new XSSFWorkbook();

// Tạo một Sheet mới trong Workbook
Sheet sheet = workbook.createSheet("Scan Results");

// Tạo các tiêu đề cho các cột trong Sheet
Row headerRow = sheet.createRow(0);
headerRow.createCell(0).setCellValue("File Name");
headerRow.createCell(1).setCellValue("Virus Name");
headerRow.createCell(2).setCellValue("Scan Result");

// Mở file chứa kết quả quét virus và đọc nó
BufferedReader br = new BufferedReader(new FileReader("scan_result.txt"));

String line;
int rowNumber = 1; // Bắt đầu ghi dữ liệu từ hàng thứ hai của Sheet

while ((line = br.readLine()) != null) {
    // Tách các giá trị trong mỗi dòng của file và lấy ra tên file, tên virus và kết quả quét
    String[] values = line.split("\\s+");
    String fileName = values[0];
    String virusName = values[1];
    String scanResult = values[2];

    // Tạo một Row mới trong Sheet và ghi dữ liệu vào các ô tương ứng
    Row row = sheet.createRow(rowNumber++);
    row.createCell(0).setCellValue(fileName);
    row.createCell(1).setCellValue(virusName);
    row.createCell(2).setCellValue(scanResult);
}

// Đóng file
br.close();

// Ghi Workbook vào file Excel
FileOutputStream fos = new FileOutputStream("scan_results.xlsx");
workbook.write(fos);
fos.close();
