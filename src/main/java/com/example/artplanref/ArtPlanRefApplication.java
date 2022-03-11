package com.example.artplanref;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.domain.EntityScan;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
public class ArtPlanRefApplication {

	public static void main(String[] args) {
		SpringApplication.run(ArtPlanRefApplication.class, args);
	}

}
