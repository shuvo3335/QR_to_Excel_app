package com.example.qrapplication;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import com.google.zxing.integration.android.IntentIntegrator;
import com.google.zxing.integration.android.IntentResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.util.HashSet;
import java.util.Set;
import android.Manifest;
import android.content.SharedPreferences;
import android.preference.PreferenceManager;
import android.content.pm.PackageManager;
import android.util.Log;
import android.view.View;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.annotation.NonNull;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import java.io.OutputStream;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import javax.xml.stream.XMLInputFactory;

public class MainActivity extends AppCompatActivity {

    private TextView resultTextView;
    private int currentRowIndex = 0; // Initialize the current row index
    private static final int DIRECTORY_REQUEST_CODE =100;
    private Set<String> scannedQR = new HashSet<>();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        resultTextView = findViewById(R.id.resultTextView);
        Button scanButton = findViewById(R.id.scanButton);
        Button fileViewBtn = findViewById(R.id.fileBtn);

        fileViewBtn.setOnClickListener(view -> {
            String folderPath = getExternalFilesDir(null)+ "/QR Scan Data";
            // Create a Uri for the folder path
            Intent intent = new Intent(Intent.ACTION_VIEW);
            Uri uri = Uri.parse(folderPath);
            intent.setDataAndType(uri, "*/*");

            if(intent.resolveActivity(getPackageManager())!=null){
                startActivity(intent);
            }
            else {
                Toast.makeText(this, "No file Manager App found", Toast.LENGTH_LONG).show();
            }
        });


        scanButton.setOnClickListener( view -> {
            startQRScanner();
        });

        /*scanButton.setOnClickListener(view -> {
            // Check and request camera and storage permissions if not granted
            if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.CAMERA)
                    != PackageManager.PERMISSION_GRANTED ||
                    ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                            != PackageManager.PERMISSION_GRANTED)
            {

                ActivityCompat.requestPermissions(MainActivity.this,
                        new String[]{Manifest.permission.CAMERA, Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        CAMERA_PERMISSION_REQUEST_CODE);
            } else {
                // Start QR code scanner
                startQRScanner();
                Toast.makeText(this, "button pressed", Toast.LENGTH_SHORT).show();

            }
        });*/
    }
    private void startQRScanner() {
        IntentIntegrator integrator = new IntentIntegrator(this);
        integrator.setOrientationLocked(true);
        integrator.setDesiredBarcodeFormats(IntentIntegrator.ALL_CODE_TYPES);
        integrator.setPrompt("QR CODE SCANNING IN PROGRESS...");
        integrator.initiateScan();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {

        IntentResult result = IntentIntegrator.parseActivityResult(requestCode, resultCode, data);
        if (result != null) {
            if (result.getContents() != null) {
                String qrData = result.getContents();
                if (!scannedQR.contains(qrData)){
                    scannedQR.add(qrData);
                    resultTextView.setText(qrData);
                    // Parse QR data and write it to an Excel file
                    writeToExcel(qrData);
                    currentRowIndex++;
                }
                else {
                    resultTextView.setText("SCAN NEW QR CODE TO GET DATA");
                    Toast.makeText(this, "THIS QR CODE IS ALREADY SCANNED", Toast.LENGTH_LONG).show();
                }

            } else {
                resultTextView.setText("No QR code data found");
                Toast.makeText(this, "No QR data found to write", Toast.LENGTH_SHORT).show();
            }
        }

        else
        {
            super.onActivityResult(requestCode, resultCode, data);
        }

    }



    private void writeToExcel(String qrData) {
        // Define the file path
        //String filePath = getFilesDir() + "/qr_data.xls";
        String currentDate = java.time.LocalDate.now().toString();
        String folderName = "QR Scan Data";
        String fileName = "qr_data" + currentDate + ".xls";

        File internalStorageDir = getExternalFilesDir(null);

        File folder = new File(internalStorageDir, folderName);
        if (!folder.exists()) {
            folder.mkdir();
            Toast.makeText(this, "directory created", Toast.LENGTH_SHORT).show();
        }
        String filePath = folder.getAbsolutePath() + File.separator + fileName;
        Workbook workbook;
        Sheet sheet;

        // Check if the file already exists
        File file = new File(filePath);

        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {

                // Disable namespace awareness for XML parsing
                workbook = WorkbookFactory.create(fis);

                // Open the existing Excel file
                sheet = workbook.getSheetAt(0); // Assuming you have only one sheet

                // Check if the QR data is already in the HashSet
                if (!scannedQR.contains(qrData)) {
                    // Find the last row with data and increment the currentRowIndex
                    int lastRowNum = sheet.getLastRowNum();
                    currentRowIndex = lastRowNum + 1;

                    // Create a new row for each QR code scan and set the data in the corresponding cells
                    Row row = sheet.createRow(currentRowIndex);

                    // Split QR data by whitespace and write to separate cells in the row
                    String[] dataParts = qrData.split(" ");
                    for (int i = 0; i < dataParts.length; i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(dataParts[i]);
                    }

                    // Add the QR data to the HashSet to mark it as scanned
                    scannedQR.add(qrData);

                    // Save the workbook to the file
                    try (FileOutputStream fos = new FileOutputStream(filePath)) {
                        workbook.write(fos);
                        resultTextView.append("\nData written to Excel file");
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                } else {
                    // Handle the case when the QR data is a duplicate
                    resultTextView.setText("SCAN NEW QR CODE TO GET DATA");
                    Toast.makeText(this, "THIS QR CODE IS ALREADY SCANNED", Toast.LENGTH_LONG).show();
                }
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }
        } else {
            // Create a new Excel workbook and sheet if the file doesn't exist
            workbook = new HSSFWorkbook();
            sheet = workbook.createSheet("QR Data");
            // Create a new row for each QR code scan and set the data in the corresponding cells
            Row row = sheet.createRow(currentRowIndex);

            // Split QR data by whitespace and write to separate cells in the row
            String[] dataParts = qrData.split(" ");
            for (int i = 0; i < dataParts.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(dataParts[i]);
            }

            // Add the QR data to the HashSet to mark it as scanned
            scannedQR.add(qrData);

            // Save the workbook to the file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                resultTextView.append("\nData written to Excel file");
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Increment the currentRowIndex
            currentRowIndex++;
        }
    }



}
