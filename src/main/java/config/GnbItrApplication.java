package config;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.springframework.context.annotation.ComponentScan;
import service.Login;

@SpringBootApplication
@ComponentScan({"service", "utility"})
public class GnbItrApplication {

    public static void main(String[] args) {
        SpringApplication.run(GnbItrApplication.class, args);

        ApplicationContext context = new AnnotationConfigApplicationContext(GnbItrApplication.class);
        Login login = context.getBean("login", Login.class);
        login.getLoginCredentials();
    }

}
