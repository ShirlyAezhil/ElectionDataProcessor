package com.example.demo;


import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ElectionDataProcessorApplication implements CommandLineRunner {

    public static void main(String[] args) {
        System.setProperty("java.awt.headless", "false");

        SpringApplication.run(ElectionDataProcessorApplication.class, args);
    }

    @Override
    public void run(String... args) {
        javax.swing.SwingUtilities.invokeLater(() -> {
            new ExcelViewerFrame().setVisible(true);
        });
    }
}

