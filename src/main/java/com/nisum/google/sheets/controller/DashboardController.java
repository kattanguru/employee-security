package com.nisum.google.sheets.controller;

import com.nisum.google.sheets.dto.GoogleSheetDTO;
import com.nisum.google.sheets.dto.GoogleSheetResponseDTO;
import com.nisum.google.sheets.service.GoogleApiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.Map;


@RestController
public class DashboardController {

    @Autowired
    private GoogleApiService googleApiService;

    @GetMapping("/getData")
    public Map<Object, Object> readDataFromGoogleSheet() throws GeneralSecurityException, IOException {
        return googleApiService.readDataFromGoogleSheet();
    }

    @PostMapping("/createSheet")
    public GoogleSheetResponseDTO createGoogleSheet(@RequestBody List<String> emailIds)
            throws GeneralSecurityException, IOException {
        return googleApiService.createSheet(emailIds);
    }

    @PostMapping("/updateSheet")
    public String updateGoogleSheet(@RequestBody GoogleSheetDTO googleSheetDTO)
            throws GeneralSecurityException, IOException {
        return googleApiService.updateSheet(googleSheetDTO);
    }
}
