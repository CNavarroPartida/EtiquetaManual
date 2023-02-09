package com.famsa.manual.controller;

import com.famsa.manual.service.EtiquetaService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

@RestController
public class EtiquetaController {

    @Autowired
    private EtiquetaService etiquetaService;

    @GetMapping("/hello")
    public String hello() {
        return "Hello World";
    }

    @PostMapping("/manual")
    public FileSystemResource generarEtiqueta(@RequestParam("file") MultipartFile file, HttpServletResponse response) throws Exception {
        return etiquetaService.generarEtiqueta(file, response);
    }
}
