package com.kiran.ExceltoJson;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.scheduling.annotation.EnableScheduling;

/**
 * Maintaining the main class to run
 */

@EnableScheduling
@SpringBootApplication
@ComponentScan("com.kiran.ExceltoJson")
public class ECJsonApp {
	public static void main(String[] args) throws Exception {
		SpringApplication.run(ECJsonApp.class, args);
	}
}
