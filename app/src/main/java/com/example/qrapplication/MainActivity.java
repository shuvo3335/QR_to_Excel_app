package com.example.qrapplication;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.util.Log;
import android.widget.Button;
import android.widget.TextView;
import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import com.google.zxing.integration.android.IntentIntegrator;
import com.google.zxing.integration.android.IntentResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class MainActivity extends AppCompatActivity {

    private static final int CAMERA_PERMISSION_REQUEST_CODE = 200;
    private TextView resultTextView;
    private int currentRowIndex = 0; // Initialize the current row index

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        resultTextView = findViewById(R.id.resultTextView);
        Button scanButton = findViewById(R.id.scanButton);

        scanButton.setOnClickListener(view -> {
            // Check and request camera and storage permissions if not granted
            if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.CAMERA)
                    != PackageManager.PERMISSION_GRANTED ||
                    ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                            != PackageManager.PERMISSION_GRANTED) {

                ActivityCompat.requestPermissions(MainActivity.this,
                        new String[]{Manifest.permission.CAMERA, Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        CAMERA_PERMISSION_REQUEST_CODE);
            } else {
                // Start QR code scanner
                startQRScanner();
            }
        });
    }

    private void startQRScanner() {
        IntentIntegrator integrator = new IntentIntegrator(this);
        integrator.initiateScan();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        IntentResult result = IntentIntegrator.parseActivityResult(requestCode, resultCode, data);
        if (result != null) {
            if (result.getContents() != null) {
                String qrData = result.getContents();
                resultTextView.setText(qrData);

                // Parse QR data and write it to an Excel file
                writeToExcel(qrData);
                currentRowIndex++;
            } else {
                resultTextView.setText("No QR code data found");
            }
        } else {
            super.onActivityResult(requestCode, resultCode, data);
        }
    }

    /*private void writeToExcel(String qrData) {
        // Define the file path
        String filePath = getExternalFilesDir(null) + "/qr_data.xlsx";

        Workbook workbook = null;
        Sheet sheet;
        try {
            File file = new File(filePath);
            if (file.exists()) {
                FileInputStream fileInputStream = new FileInputStream(file);
                workbook = new XSSFWorkbook(fileInputStream);
                sheet = workbook.getSheetAt(0);
                int lastRowNum = sheet.getLastRowNum();
                currentRowIndex = lastRowNum + 1;
                fileInputStream.close();
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("QR DATA");
            }
            Row row = sheet.createRow(currentRowIndex);
            String[] dataParts = qrData.split(" ");
            for (int i = 0; i < dataParts.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(dataParts[i]);
            }
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);
            resultTextView.append("\nData written to Excel file");
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }*/


    private void writeToExcel(String qrData) {
        // Define the file path
        String filePath = getExternalFilesDir(null) + "/qr_dataa.xlsx";

        XSSFWorkbook workbook;
        XSSFSheet sheet;

        // Check if the file already exists
        File file = new File(filePath);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                // Open the existing Excel file
                workbook = new XSSFWorkbook(fis);
                sheet = workbook.getSheetAt(0); // Assuming you have only one sheet

                // Find the last row with data and increment the currentRowIndex
                int lastRowNum = sheet.getLastRowNum();
                currentRowIndex = lastRowNum + 1;
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }
        } else {
            // Create a new Excel workbook and sheet if the file doesn't exist
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("QR Data");
        }

        // Create a new row for each QR code scan and set the data in the corresponding cells
        XSSFRow row = sheet.createRow(currentRowIndex);

        // Split QR data by whitespace and write to separate cells in the row
        String[] dataParts = qrData.split(" ");
        for (int i = 0; i < dataParts.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(dataParts[i]);
        }

        // Save the workbook to the file
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            resultTextView.append("\nData written to Excel file");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
