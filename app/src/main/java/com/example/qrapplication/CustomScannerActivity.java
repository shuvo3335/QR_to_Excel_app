package com.example.qrapplication;

import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import com.google.zxing.Result;
import com.google.zxing.ResultPoint;
import com.journeyapps.barcodescanner.BarcodeCallback;
import com.journeyapps.barcodescanner.BarcodeResult;
import com.journeyapps.barcodescanner.CompoundBarcodeView;

import java.util.List;

public class CustomScannerActivity extends AppCompatActivity {

    private CompoundBarcodeView barcodeView;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_custom_scanner);

        barcodeView = findViewById(R.id.barcode_scanner);

        // Set up a callback for handling scanned QR codes
        barcodeView.decodeContinuous(new BarcodeCallback() {
            @Override
            public void barcodeResult(BarcodeResult result) {
                if (result != null) {
                    String qrData = result.getText();
                    Log.d("QRCodeScanner", "Scanned QR code: " + qrData);

                    // Handle the scanned QR data as needed (e.g., pass it to MainActivity)
                    // You can use Intent or other mechanisms to send data back to the calling activity
                    Intent intent = new Intent();
                    intent.putExtra("QR_DATA", qrData);
                    setResult(RESULT_OK, intent);
                    finish();
                }
            }

            @Override
            public void possibleResultPoints(List<ResultPoint> resultPoints) {
                // Optional: Handle possible result points for advanced features
            }
        });
    }

    @Override
    protected void onResume() {
        super.onResume();
        // Start the camera preview when the activity resumes
        barcodeView.resume();
    }

    @Override
    protected void onPause() {
        super.onPause();
        // Pause the camera preview when the activity is paused
        barcodeView.pause();
    }
}
