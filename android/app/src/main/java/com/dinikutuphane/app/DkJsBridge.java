package com.dinikutuphane.app;

import android.app.Activity;
import android.webkit.JavascriptInterface;

public class DkJsBridge {
    private final Activity activity;

    public DkJsBridge(Activity activity) {
        this.activity = activity;
    }

    @JavascriptInterface
    public void exit() {
        activity.runOnUiThread(new Runnable() {
            @Override
            public void run() {
                activity.finishAffinity();
            }
        });
    }
}