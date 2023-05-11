package utility;

import org.springframework.context.annotation.Configuration;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

@Configuration
public class ConfigUtility {
    public String getProperty(String propertyKey) {
        try {
            Properties properties = new Properties();
            FileInputStream propertyFile = new FileInputStream(System.getProperty("user.dir") + "\\src/main/resources/application.properties");
            properties.load(propertyFile);
            return properties.getProperty(propertyKey);
        } catch (FileNotFoundException exception) {
            throw new RuntimeException(exception);
        } catch (IOException exception) {
            throw new RuntimeException(exception);
        }
    }

}
