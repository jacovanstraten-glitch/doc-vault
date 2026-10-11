package io.github.jacovanstraten.vault;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.MediaStore;
import android.view.View;
import android.view.WindowInsets;
import android.webkit.CookieManager;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.FrameLayout;
import android.widget.Toast;

import androidx.core.content.FileProvider;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Document Vault for Android: a full-screen window around the Document Vault web app.
 * The app itself (and every update) comes from GitHub Pages; documents stay in OneDrive.
 */
public class MainActivity extends Activity {

    private static final String APP_URL = "https://jacovanstraten-glitch.github.io/doc-vault/";
    private static final String APP_HOST = "jacovanstraten-glitch.github.io";
    private static final String APP_PATH = "/doc-vault";
    private static final int REQ_FILES = 41;

    private WebView web;
    private ValueCallback<Uri[]> fileCallback;
    private Uri cameraUri;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        FrameLayout root = new FrameLayout(this);
        root.setBackgroundColor(Color.parseColor("#1E6A55"));
        web = new WebView(this);
        web.setBackgroundColor(Color.parseColor("#F2F5F3"));
        root.addView(web, new FrameLayout.LayoutParams(
                FrameLayout.LayoutParams.MATCH_PARENT, FrameLayout.LayoutParams.MATCH_PARENT));
        setContentView(root);

        // Keep the page clear of the status bar, navigation bar and keyboard (edge-to-edge on Android 15+).
        root.setOnApplyWindowInsetsListener((View v, WindowInsets insets) -> {
            if (Build.VERSION.SDK_INT >= 30) {
                android.graphics.Insets b = insets.getInsets(WindowInsets.Type.systemBars() | WindowInsets.Type.ime());
                v.setPadding(b.left, b.top, b.right, b.bottom);
                return WindowInsets.CONSUMED;
            }
            v.setPadding(insets.getSystemWindowInsetLeft(), insets.getSystemWindowInsetTop(),
                    insets.getSystemWindowInsetRight(), insets.getSystemWindowInsetBottom());
            return insets.consumeSystemWindowInsets();
        });

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setAllowFileAccess(false);
        s.setAllowContentAccess(true);
        s.setSupportMultipleWindows(false);
        s.setMediaPlaybackRequiresUserGesture(true);

        CookieManager cookies = CookieManager.getInstance();
        cookies.setAcceptCookie(true);
        cookies.setAcceptThirdPartyCookies(web, true);

        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                Uri u = request.getUrl();
                if (staysInApp(u)) return false;
                openOutside(u);
                return true;
            }
        });

        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> callback, FileChooserParams params) {
                return chooseFiles(callback, params);
            }
        });

        if (savedInstanceState != null) {
            web.restoreState(savedInstanceState);
        } else {
            web.loadUrl(APP_URL);
        }
    }

    /** The app's own pages and Microsoft sign-in stay inside; everything else (OneDrive web, links) opens outside. */
    private static boolean staysInApp(Uri u) {
        String scheme = u.getScheme();
        String host = u.getHost() == null ? "" : u.getHost().toLowerCase();
        if (!"https".equals(scheme)) return false;
        if (host.equals(APP_HOST)) return u.getPath() != null && u.getPath().startsWith(APP_PATH);
        if (host.startsWith("onedrive.") || host.contains("1drv")) return false;
        return host.endsWith("login.microsoftonline.com") || host.endsWith("login.live.com")
                || host.endsWith("account.live.com") || host.endsWith("signup.live.com")
                || host.endsWith("login.microsoft.com") || host.endsWith("account.microsoft.com")
                || host.endsWith("msauth.net") || host.endsWith("msftauth.net");
    }

    private void openOutside(Uri u) {
        try {
            startActivity(new Intent(Intent.ACTION_VIEW, u));
        } catch (ActivityNotFoundException e) {
            Toast.makeText(this, "No app found to open this link.", Toast.LENGTH_SHORT).show();
        }
    }

    /** Camera for "Scan document" (capture), the system file picker for everything else. */
    private boolean chooseFiles(ValueCallback<Uri[]> callback, WebChromeClient.FileChooserParams params) {
        if (fileCallback != null) fileCallback.onReceiveValue(null);
        fileCallback = callback;
        cameraUri = null;

        List<String> mimes = new ArrayList<>();
        boolean imagesOnly = true;
        for (String raw : params.getAcceptTypes()) {
            for (String t : raw.split(",")) {
                t = t.trim().toLowerCase();
                if (t.isEmpty()) continue;
                if (t.equals(".pdf")) t = "application/pdf";
                if (!t.contains("/")) continue;
                mimes.add(t);
                if (!t.startsWith("image/")) imagesOnly = false;
            }
        }

        try {
            if (params.isCaptureEnabled() && imagesOnly) {
                File dir = new File(getCacheDir(), "camera");
                //noinspection ResultOfMethodCallIgnored
                dir.mkdirs();
                File photo = new File(dir, "scan-" + System.currentTimeMillis() + ".jpg");
                cameraUri = FileProvider.getUriForFile(this, getPackageName() + ".files", photo);
                Intent cam = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
                cam.putExtra(MediaStore.EXTRA_OUTPUT, cameraUri);
                cam.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_READ_URI_PERMISSION);
                startActivityForResult(cam, REQ_FILES);
                return true;
            }
            Intent pick = new Intent(Intent.ACTION_GET_CONTENT);
            pick.addCategory(Intent.CATEGORY_OPENABLE);
            pick.setType(mimes.size() == 1 ? mimes.get(0) : "*/*");
            if (mimes.size() > 1) pick.putExtra(Intent.EXTRA_MIME_TYPES, mimes.toArray(new String[0]));
            pick.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, params.getMode() == WebChromeClient.FileChooserParams.MODE_OPEN_MULTIPLE);
            startActivityForResult(Intent.createChooser(pick, "Choose files"), REQ_FILES);
            return true;
        } catch (Exception e) {
            fileCallback = null;
            cameraUri = null;
            Toast.makeText(this, "Could not open the camera or file picker.", Toast.LENGTH_SHORT).show();
            return false;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode != REQ_FILES) {
            super.onActivityResult(requestCode, resultCode, data);
            return;
        }
        Uri[] result = null;
        if (resultCode == RESULT_OK) {
            if (data != null && data.getClipData() != null) {
                ClipData clip = data.getClipData();
                result = new Uri[clip.getItemCount()];
                for (int i = 0; i < clip.getItemCount(); i++) result[i] = clip.getItemAt(i).getUri();
            } else if (data != null && data.getData() != null) {
                result = new Uri[]{data.getData()};
            } else if (cameraUri != null) {
                result = new Uri[]{cameraUri};
            }
        }
        if (fileCallback != null) fileCallback.onReceiveValue(result);
        fileCallback = null;
        cameraUri = null;
    }

    /** Back closes an open panel in the app first, then goes back, then leaves the app. */
    @Override
    public void onBackPressed() {
        String js = "(function(){var p=document.getElementById('panel');"
                + "if(p&&!p.hidden){var b=p.querySelector('[data-act=\"close\"]');if(b){b.click();return true;}}"
                + "return false;})()";
        web.evaluateJavascript(js, handled -> {
            if ("true".equals(handled)) return;
            if (web.canGoBack()) web.goBack();
            else finish();
        });
    }

    @Override
    protected void onSaveInstanceState(Bundle outState) {
        super.onSaveInstanceState(outState);
        web.saveState(outState);
    }

    @Override
    protected void onPause() {
        super.onPause();
        web.onPause();
        CookieManager.getInstance().flush();
    }

    @Override
    protected void onResume() {
        super.onResume();
        web.onResume();
    }

    @Override
    protected void onDestroy() {
        if (web != null) web.destroy();
        super.onDestroy();
    }
}
