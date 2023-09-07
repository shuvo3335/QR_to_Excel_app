package com.example.qrapplication;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
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

import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    private static final int CAMERA_PERMISSION_REQUEST_CODE = 200;
    private TextView resultTextView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        resultTextView = findViewById(R.id.resultTextView);
        Button scanButton = findViewById(R.id.scanButton);

        scanButton.setOnClickListener(view -> {
            // Check and request camera permission if not granted
            if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.CAMERA)
                    != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(MainActivity.this,
                        new String[]{Manifest.permission.CAMERA}, CAMERA_PERMISSION_REQUEST_CODE);
            } else {
                // Start QR code scanner
                startQRScanner();
            }
        });
    }

    private void startQRScanner() {
        IntentIntegrator integrator = new IntentIntegrator(this);
        //integrator.setCaptureActivity(CustomScannerActivity.class); // Custom QR code scanner activity
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
            } else {
                resultTextView.setText("No QR code data found");
            }
        } else {
            super.onActivityResult(requestCode, resultCode, data);
        }
    }

    private void writeToExcel(String qrData) {
        // Create a new Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("QR Data");

        // Split QR data by whitespace and write to separate columns
        String[] dataParts = qrData.split(" ");
        Row row = sheet.createRow(0);
        for (int i = 0; i < dataParts.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(dataParts[i]);
        }

        // Save the workbook to a file
        try {
            FileOutputStream fos = new FileOutputStream(getExternalFilesDir(null) + "/qr_data.xlsx");
            workbook.write(fos);
            fos.close();
            resultTextView.append("\nData written to Excel file");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
