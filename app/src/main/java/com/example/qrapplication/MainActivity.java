package com.example.qrapplication;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.util.Log;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import com.google.zxing.integration.android.IntentIntegrator;
import com.google.zxing.integration.android.IntentResult;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.xml.stream.XMLInputFactory;

public class MainActivity extends AppCompatActivity {

    private static final int CAMERA_PERMISSION_REQUEST_CODE = 200;
    private TextView resultTextView;
    private int currentRowIndex = 0, num = 0; // Initialize the current row index

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
                resultTextView.setText(qrData);

                // Parse QR data and write it to an Excel file
                writeToExcel(qrData);
                currentRowIndex++;
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
        String filePath = getExternalFilesDir(null) + "/qr_data.xls";

        Workbook workbook;
        Sheet sheet;

        // Check if the file already exists
        File file = new File(filePath);

        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {

                //Disable namespace awareness for XML parsing
                /*XMLInputFactory factory = XMLInputFactory.newInstance();
                factory.setProperty(XMLInputFactory.IS_NAMESPACE_AWARE, false);*/
                workbook = WorkbookFactory.create(fis);

                // Open the existing Excel file
                Log.d("File exist tag", "writeToExcel: file existing error");
//                workbook = new XSSFWorkbook();
                sheet = workbook.getSheetAt(0); // Assuming you have only one sheet

                // Find the last row with data and increment the currentRowIndex
                int lastRowNum = sheet.getLastRowNum();
                currentRowIndex = lastRowNum+1;
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }
        }

        else {
            // Create a new Excel workbook and sheet if the file doesn't exist
            workbook = new HSSFWorkbook();
            sheet = workbook.createSheet("QR Data"); //+" "+ num++
        }

        // Create a new row for each QR code scan and set the data in the corresponding cells
        Row row = sheet.createRow(currentRowIndex);

        // Split QR data by whitespace and write to separate cells in the row
        String[] dataParts = qrData.split(" ");
        for (int i = 0; i < dataParts.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(dataParts[i]);
        }

            /*try (OutputStream outputStream = new FileOutputStream(filePath))
            {
                workbook.write(outputStream);
                Toast.makeText(this, "error occure here", Toast.LENGTH_SHORT).show();
                // Create a new row for each QR code scan and set the data in the corresponding cells
                Row row = sheet.createRow(2);
                currentRowIndex++;
                // Split QR data by whitespace and write to separate cells in the row
                String[] dataParts = qrData.split(" ");
                for (int i = 0; i < dataParts.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(dataParts[i]);
            }
        } catch (IOException e) {
                throw new RuntimeException(e);
            }*/


        // Save the workbook to the file
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            resultTextView.append("\nData written to Excel file");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
