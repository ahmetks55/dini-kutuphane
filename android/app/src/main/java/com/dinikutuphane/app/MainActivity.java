package com.dinikutuphane.app;

import android.os.Bundle;
import android.webkit.WebView;

import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        WebView wv = getBridge().getWebView();
        if (wv != null) {
            wv.addJavascriptInterface(new DkJsBridge(this), "dkBridge");
        }
    }

    @Override
    public void onBackPressed() {
        WebView wv = getBridge().getWebView();
        if (wv != null && wv.canGoBack()) {
            wv.goBack();
            return;
        }
        if (wv != null) {
            wv.evaluateJavascript("window.dispatchEvent(new CustomEvent('dknativeback', {}));", value -> {});
            return;
        }
        super.onBackPressed();
    }
}