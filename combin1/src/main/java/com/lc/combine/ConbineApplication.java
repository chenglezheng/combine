package com.lc.combine;

import com.lc.combine.service.CombineExcelService;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import static org.springframework.boot.SpringApplication.run;

@SpringBootApplication
public class ConbineApplication {

	public static void main(String[] args) {
		ConfigurableApplicationContext context = SpringApplication.run(ConbineApplication.class, args);
		CombineExcelService combineExcelService= context.getBean(CombineExcelService.class);
		try{
			combineExcelService.combinne();
		}catch(Exception e){
		    e.printStackTrace();
		}
	}

}
