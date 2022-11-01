package ru.p03;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ApplicationProperties {
    private final Properties configProp = new Properties();

    private ApplicationProperties() {
        File f = new File("./conf/application.properties");
        String s = null;
        try {
            s = f.getCanonicalPath();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Properties: " + s);
        InputStream is = null;
        try {
            is = new FileInputStream(s);
            configProp.load(is);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static class LazyHolder {
        private static final ApplicationProperties INSTANCE = new ApplicationProperties();
    }

    public static ApplicationProperties getInstance() {
        return LazyHolder.INSTANCE;
    }

    public String getProperty(String key) {
        return configProp.getProperty(key);
    }

    public Properties getConfigProps() {
        return configProp;
    }
}