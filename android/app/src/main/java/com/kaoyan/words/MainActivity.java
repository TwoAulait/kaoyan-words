package com.kaoyan.words;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.Locale;

/**
 * 考研英语背单词（安卓端）
 * WebView 外壳加载 assets/index.html；addJavascriptInterface 提供
 * 备份文件导出/导入（走系统文件选择器 SAF）与 toast；
 * 坚果云 WebDAV 请求走原生 HttpURLConnection 桥（绕开 CORS）。
 */
public class MainActivity extends Activity {

    private static final int REQ_EXPORT = 1001;
    private static final int REQ_IMPORT = 1002;

    private WebView webView;
    private Bridge bridge;
    private String pendingExportJson = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        webView = new WebView(this);
        setContentView(webView);

        WebSettings s = webView.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setAllowFileAccess(true);
        s.setAllowContentAccess(true);

        webView.setWebViewClient(new WebViewClient());
        bridge = new Bridge();
        webView.addJavascriptInterface(bridge, "AndroidBridge");
        webView.loadUrl("file:///android_asset/index.html");
    }

    @Override
    public void onBackPressed() {
        // 手动同步：退出时若有未上传到云端的改动，弹窗三选（先上传再退出/直接退出/继续学习）。
        // 必须由原生处理——退出的瞬间 JS 不可靠，上传也用原生直接发请求。
        if (bridge != null && bridge.isCloudConfigured() && bridge.dirty) {
            AlertDialog.Builder b = new AlertDialog.Builder(this);
            b.setTitle("退出确认");
            b.setMessage("本机还有改动未上传到云端。\n\n"
                    + "「先上传再退出」：上传到云端后退出；\n"
                    + "「直接退出」：不上传，本次改动只保存在本地备份里；\n"
                    + "「继续学习」：留在本页，稍后再传。");
            b.setPositiveButton("先上传再退出", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface d, int w) { bridge.uploadAndExit(); }
            });
            b.setNegativeButton("直接退出", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface d, int w) { finish(); }
            });
            b.setNeutralButton("继续学习", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface d, int w) { /* 留在本页 */ }
            });
            b.setCancelable(false);
            b.show();
            return;
        }
        finish();
    }

    @Override
    protected void onPause() {
        super.onPause();
        // 切后台/锁屏：有未上传改动时静默做一份本地备份（不联网、不提示）。
        // 上传只发生在用户手动点按钮，或退出时选择「先上传再退出」。
        if (bridge == null) return;
        long now = System.currentTimeMillis();
        if (now - lastExitBackupMs < 1000) return;   // onPause 可能被多次触发，去重
        lastExitBackupMs = now;
        if (bridge.dirty) bridge.writeBackupFile(bridge.buildBackupJson());
    }

    private long lastExitBackupMs = 0;

    // ------------------------------------------------------------------ JS 桥
    private class Bridge {
        // —— 手机端状态快照与云端配置：JS 在每次保存/配置变更时推给原生，
        //    原生在 onPause / 退出时用快照直接「写本地备份 + 上传云端」，
        //    不依赖 JS 在后台能否执行（WebView 后台冻结/删进程时 JS 不可靠）。
        volatile int lastProgress = -1;
        volatile String lastUnknownJson = null;
        volatile boolean dirty = false;             // JS 推来：本机相对云端是否有未上传改动
        volatile String cloudUrl = null, cloudUser = null, cloudPass = null;
        private static final int BACKUP_KEEP = 10;

        @JavascriptInterface
        public void updateState(String progress, String unknownJson) {
            try { lastProgress = Integer.parseInt(progress); } catch (Exception ignored) { }
            lastUnknownJson = unknownJson;
        }

        @JavascriptInterface
        public void setCloudConfig(String url, String user, String pass) {
            cloudUrl = emptyToNull(url);
            cloudUser = emptyToNull(user);
            cloudPass = emptyToNull(pass);
        }

        @JavascriptInterface
        public void setDirty(boolean d) {
            dirty = d;
        }

        public boolean isCloudConfigured() {
            return cloudUrl != null && cloudUser != null && cloudPass != null;
        }

        private static String emptyToNull(String s) {
            return (s == null || s.trim().isEmpty()) ? null : s.trim();
        }

        // —— 原生本地备份：存到 getFilesDir()/本地备份/（进程被杀也持久），
        //    设置页列表经 listLocalBackups / readLocalBackup 异步读取回 JS。
        @JavascriptInterface
        public void saveLocalBackup(String backupJson) {
            writeBackupFile(backupJson);
        }

        @JavascriptInterface
        public void listLocalBackups() {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    final String json = listBackupsJson();
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() { callBridge("onLocalBackups", json); }
                    });
                }
            }).start();
        }

        @JavascriptInterface
        public void readLocalBackup(int index) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    final String json = readBackupAt(index);
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() { callBridge("onLocalBackupRead", json); }
                    });
                }
            }).start();
        }

        /** 「先上传再退出」：先写本地备份，再读云端二次确认（显示云端与本机数字），确认后上传；成功退出，失败提示并留在本页。 */
        public void uploadAndExit() {
            if (lastProgress < 0 || lastUnknownJson == null) { finish(); return; }
            final String json = buildBackupJson();
            writeBackupFile(json);   // 本地备份同步落盘，先保证数据安全
            if (!isCloudConfigured()) { finish(); return; }
            final String u = cloudUrl, us = cloudUser, ps = cloudPass;
            new Thread(new Runnable() {
                @Override
                public void run() {
                    final String cloudDesc = readCloudDesc(u, us, ps);
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() { confirmUploadThenExit(cloudDesc, json, u, us, ps); }
                    });
                }
            }).start();
        }

        /** 读云端当前数据，描述成「第 X / 5493 词，生词 N 个」；读不到返回原因。 */
        private String readCloudDesc(String u, String us, String ps) {
            try {
                JSONObject r = new JSONObject(httpJson("GET", u, us, ps, null));
                if (!r.optBoolean("ok", false)) return "云端读取失败，未知";
                int st = r.optInt("status", -1);
                if (st == 404) return "暂无数据";
                if (st != 200) return "云端读取失败（HTTP " + st + "）";
                JSONObject d = new JSONObject(r.optString("body", "{}"));
                int p = d.optInt("progress", 0);
                JSONArray rows = d.optJSONArray("unknown");
                int n = rows == null ? 0 : rows.length();
                if (p <= 0 && n <= 0) return "暂无数据";
                return "第 " + (p + 1) + " / 5493 词，生词 " + n + " 个";
            } catch (Exception e) {
                return "云端读取失败";
            }
        }

        /** 二次确认弹窗：显示云端与本机数字，确认才上传并退出。 */
        private void confirmUploadThenExit(final String cloudDesc, final String json,
                                           final String u, final String us, final String ps) {
            AlertDialog.Builder b = new AlertDialog.Builder(MainActivity.this);
            b.setTitle("上传确认");
            b.setMessage("云端当前：" + cloudDesc + "\n本机当前：第 " + (lastProgress + 1)
                    + " / 5493 词，生词 " + unknownCount() + " 个\n\n确认把本机数据上传覆盖云端吗？");
            b.setPositiveButton("确认上传", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface d, int w) { doUploadAndExit(json, u, us, ps); }
            });
            b.setNegativeButton("取消", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface d, int w) { /* 留在本页，稍后再传 */ }
            });
            b.setCancelable(false);
            b.show();
        }

        private int unknownCount() {
            try {
                if (lastUnknownJson == null || lastUnknownJson.trim().isEmpty()) return 0;
                return new JSONArray(lastUnknownJson).length();
            } catch (Exception e) { return 0; }
        }

        /** 上传并退出：成功 finish()，失败提示并留在本页。 */
        private void doUploadAndExit(final String body, final String u, final String us, final String ps) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    boolean ok = false;
                    try {
                        JSONObject r = new JSONObject(httpJson("PUT", u, us, ps, body));
                        int st = r.optInt("status", -1);
                        ok = r.optBoolean("ok", false) && st >= 200 && st < 300;
                    } catch (Exception ignored) { }
                    if (ok) {
                        runOnUiThread(new Runnable() {
                            @Override
                            public void run() { finish(); }
                        });
                    } else {
                        runOnUiThread(new Runnable() {
                            @Override
                            public void run() {
                                Toast.makeText(getApplicationContext(),
                                        "云端上传失败，本次未退出", Toast.LENGTH_SHORT).show();
                            }
                        });
                    }
                }
            }).start();
        }

        private String buildBackupJson() {
            JSONObject o = new JSONObject();
            try {
                o.put("version", 1);
                o.put("progress", lastProgress);
                o.put("unknown", (lastUnknownJson == null || lastUnknownJson.trim().isEmpty())
                        ? new JSONArray() : new JSONArray(lastUnknownJson));
                o.put("backup_time", new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.US).format(new Date()));
            } catch (Exception ignored) { }
            return o.toString();
        }

        private File backupsDir() {
            File d = new File(MainActivity.this.getFilesDir(), "本地备份");
            if (!d.exists()) d.mkdirs();
            return d;
        }

        /** 备份文件列表，按修改时间新→旧。 */
        private File[] backupFiles() {
            File[] all = backupsDir().listFiles(new FilenameFilter() {
                @Override
                public boolean accept(File dir, String name) {
                    return name.startsWith("备份_") && name.endsWith(".json");
                }
            });
            if (all == null) return new File[0];
            Arrays.sort(all, new Comparator<File>() {
                @Override
                public int compare(File a, File b) {
                    return Long.compare(b.lastModified(), a.lastModified());
                }
            });
            return all;
        }

        private void writeBackupFile(String backupJson) {
            try {
                File d = backupsDir();
                // 与电脑端命名一致：备份_时间戳[序号].json，保留最近 10 条
                String base = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS", Locale.US).format(new Date());
                File f = new File(d, "备份_" + base + ".json");
                int n = 1;
                while (f.exists()) {
                    f = new File(d, "备份_" + base + "_" + n + ".json");
                    n++;
                }
                FileOutputStream os = new FileOutputStream(f);
                os.write(backupJson.getBytes("UTF-8"));
                os.flush();
                os.close();
                File[] all = backupFiles();
                for (int i = BACKUP_KEEP; i < all.length; i++) {
                    all[i].delete();
                }
            } catch (Exception ignored) { }
        }

        private String readFile(File f) {
            try {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                InputStream is = new java.io.FileInputStream(f);
                byte[] buf = new byte[8192];
                int n;
                while ((n = is.read(buf)) != -1) {
                    bos.write(buf, 0, n);
                }
                is.close();
                return new String(bos.toByteArray(), "UTF-8");
            } catch (Exception e) {
                return "";
            }
        }

        private String listBackupsJson() {
            JSONArray arr = new JSONArray();
            for (File f : backupFiles()) {
                try {
                    JSONObject o = new JSONObject();
                    o.put("name", f.getName());
                    o.put("t", f.lastModified());
                    JSONObject d = new JSONObject(readFile(f));
                    o.put("progress", d.optInt("progress", 0));
                    JSONArray unk = d.optJSONArray("unknown");
                    o.put("unknown_total", unk == null ? 0 : unk.length());
                    arr.put(o);
                } catch (Exception ignored) { }
            }
            return arr.toString();
        }

        private String readBackupAt(int index) {
            File[] all = backupFiles();
            if (index < 0 || index >= all.length) return "null";
            return readFile(all[index]);
        }

        @JavascriptInterface
        public void exportBackup(String json) {
            pendingExportJson = json;
            Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            i.addCategory(Intent.CATEGORY_OPENABLE);
            i.setType("application/json");
            i.putExtra(Intent.EXTRA_TITLE, "考研英语背单词备份.json");
            startActivityForResult(i, REQ_EXPORT);
        }

        @JavascriptInterface
        public void importBackup() {
            Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            i.addCategory(Intent.CATEGORY_OPENABLE);
            i.setType("application/json");
            startActivityForResult(i, REQ_IMPORT);
        }

        @JavascriptInterface
        public void toast(String msg) {
            MainActivity.this.runOnUiThread(new Runnable() {
                @Override
                public void run() {
                    Toast.makeText(MainActivity.this, msg == null ? "" : msg, Toast.LENGTH_SHORT).show();
                }
            });
        }

        // 坚果云 WebDAV：用原生 HttpURLConnection + Basic 认证，绕开 WebView fetch 的 CORS 限制。
        // 请求在后台线程完成，结果经 window.BridgeCallbacks.onCloud(cbId, resultJson) 回传 JS。
        // resultJson 形如 {"ok":true,"status":200,"body":"..."} 或 {"ok":false,"err":"..."}。
        @JavascriptInterface
        public void cloudGet(String url, String user, String pass, String cbId) {
            doCloudRequest("GET", url, user, pass, null, cbId);
        }

        @JavascriptInterface
        public void cloudPut(String url, String user, String pass, String body, String cbId) {
            doCloudRequest("PUT", url, user, pass, body, cbId);
        }

        private void doCloudRequest(final String method, final String url, final String user,
                                    final String pass, final String body, final String cbId) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    final String result = httpJson(method, url, user, pass, body);
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() {
                            // JSONObject.quote 把回调名与结果都转成合法的 JS 字符串字面量
                            webView.evaluateJavascript(
                                "window.BridgeCallbacks && window.BridgeCallbacks.onCloud && "
                                + "window.BridgeCallbacks.onCloud(" + JSONObject.quote(cbId) + ", "
                                + JSONObject.quote(result) + ");", null);
                        }
                    });
                }
            }).start();
        }

        private String httpJson(String method, String url, String user, String pass, String body) {
            java.net.HttpURLConnection conn = null;
            try {
                conn = (java.net.HttpURLConnection) new java.net.URL(url).openConnection();
                conn.setRequestMethod(method);
                conn.setConnectTimeout(15000);
                conn.setReadTimeout(15000);
                String auth = android.util.Base64.encodeToString(
                    (user + ":" + pass).getBytes("UTF-8"), android.util.Base64.NO_WRAP);
                conn.setRequestProperty("Authorization", "Basic " + auth);
                if (body != null) {
                    conn.setDoOutput(true);
                    conn.setRequestProperty("Content-Type", "application/json");
                    conn.getOutputStream().write(body.getBytes("UTF-8"));
                }
                int code = conn.getResponseCode();
                InputStream is = (code >= 200 && code < 300) ? conn.getInputStream() : conn.getErrorStream();
                String content = "";
                if (is != null) {
                    ByteArrayOutputStream bos = new ByteArrayOutputStream();
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = is.read(buf)) != -1) {
                        bos.write(buf, 0, n);
                    }
                    content = new String(bos.toByteArray(), "UTF-8");
                }
                return new JSONObject().put("ok", true).put("status", code).put("body", content).toString();
            } catch (Exception e) {
                try {
                    return new JSONObject().put("ok", false)
                            .put("err", String.valueOf(e.getMessage())).toString();
                } catch (Exception e2) {
                    return "{\"ok\":false,\"err\":\"unknown\"}";
                }
            } finally {
                if (conn != null) {
                    conn.disconnect();
                }
            }
        }
    }

    // ------------------------------------------------------------------ SAF 回调
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (resultCode != RESULT_OK || data == null || data.getData() == null) {
            return;
        }
        final Uri uri = data.getData();
        if (requestCode == REQ_EXPORT) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    writeBackup(uri);
                }
            }).start();
        } else if (requestCode == REQ_IMPORT) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    readBackup(uri);
                }
            }).start();
        }
    }

    private void writeBackup(Uri uri) {
        try {
            OutputStream os = getContentResolver().openOutputStream(uri);
            if (os == null) {
                callBridge("onExportError", "无法写入所选位置");
                return;
            }
            os.write(pendingExportJson == null ? "".getBytes("UTF-8") : pendingExportJson.getBytes("UTF-8"));
            os.flush();
            os.close();
            callBridge("onExportDone", "备份已导出");
        } catch (Exception e) {
            callBridge("onExportError", "导出失败：" + e.getMessage());
        }
    }

    private void readBackup(Uri uri) {
        try {
            InputStream is = getContentResolver().openInputStream(uri);
            if (is == null) {
                callBridge("onImportError", "无法读取所选文件");
                return;
            }
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            byte[] buf = new byte[8192];
            int n;
            while ((n = is.read(buf)) != -1) {
                bos.write(buf, 0, n);
            }
            is.close();
            callBridge("onImportDone", new String(bos.toByteArray(), "UTF-8"));
        } catch (Exception e) {
            callBridge("onImportError", "导入失败：" + e.getMessage());
        }
    }

    private void callBridge(final String fn, final String arg) {
        runOnUiThread(new Runnable() {
            @Override
            public void run() {
                // JSONObject.quote 把字符串转义为合法 JSON 字符串字面量，防注入
                String safe = JSONObject.quote(arg == null ? "" : arg);
                webView.evaluateJavascript("window.BridgeCallbacks." + fn + "(" + safe + ");", null);
            }
        });
    }
}
