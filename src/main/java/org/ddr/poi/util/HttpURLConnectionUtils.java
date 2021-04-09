package org.ddr.poi.util;

import org.apache.commons.lang3.StringUtils;

import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.SSLSocketFactory;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.KeyManagementException;
import java.security.NoSuchAlgorithmException;
import java.security.NoSuchProviderException;
import java.security.SecureRandom;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.util.Base64;

/**
 * HttpURLConnection工具类，信任所有https地址
 *
 * @author Draco
 * @since 2019-12-12
 */
public class HttpURLConnectionUtils {
    public static class X509TrustAllManager implements X509TrustManager {
        @Override
        public void checkClientTrusted(X509Certificate[] x509Certificates, String s) throws CertificateException {
        }

        @Override
        public void checkServerTrusted(X509Certificate[] x509Certificates, String s) throws CertificateException {
        }

        @Override
        public X509Certificate[] getAcceptedIssuers() {
            return null;
        }
    }

    public static class TrustAllHostname implements HostnameVerifier {
        @Override
        public boolean verify(String s, SSLSession sslSession) {
            return true;
        }
    }

    /**
     * 开启http连接，默认未开启doOutput
     *
     * @param urlSpec url
     * @return http连接
     */
    public static HttpURLConnection connect(String urlSpec) throws IOException {
        return connect(urlSpec, null, null);
    }

    /**
     * 开启http连接，默认未开启doOutput
     *
     * @param urlSpec url
     * @param user basic auth用户名
     * @param password basic auth密码
     * @return http连接
     */
    public static HttpURLConnection connect(String urlSpec, String user, String password) throws IOException {
        if (!StringUtils.startsWith(urlSpec, "http")) {
            throw new IllegalArgumentException("Illegal url: " + urlSpec);
        }

        URL url = new URL(urlSpec);
        HttpURLConnection httpURLConnection = (HttpURLConnection) url.openConnection();
        httpURLConnection.setUseCaches(false);

        boolean isHttps = StringUtils.startsWith(urlSpec, "https");
        if (isHttps) {
            HttpsURLConnection httpsURLConnection = (HttpsURLConnection) httpURLConnection;
            httpsURLConnection.setSSLSocketFactory(trustAllSslSocketFactory());
            httpsURLConnection.setHostnameVerifier(new TrustAllHostname());
        }

        if (user != null) {
            if (password == null) {
                password = "";
            }
            String credential = user + ":" + password;
            String auth = "Basic " + Base64.getEncoder().encodeToString(credential.getBytes());
            httpURLConnection.setRequestProperty("Authorization", auth);
        }

        return httpURLConnection;
    }

    public static SSLSocketFactory trustAllSslSocketFactory() {
        try {
            TrustManager[] trustManagers = {new X509TrustAllManager()};
            SSLContext sslContext = SSLContext.getInstance("SSL", "SunJSSE");
            sslContext.init(null, trustManagers, SecureRandom.getInstance("SHA1PRNG"));
            return sslContext.getSocketFactory();
        } catch (NoSuchAlgorithmException | NoSuchProviderException | KeyManagementException e) {
            throw new RuntimeException(e);
        }
    }
}
