package com.nisum.google.sheets.service;

import com.nisum.google.sheets.dto.GoogleSheetDTO;
import com.nisum.google.sheets.dto.GoogleSheetResponseDTO;
import com.nisum.google.sheets.util.GoogleApiUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.Map;

@Service
public class GoogleApiService {

    @Autowired
    private GoogleApiUtil googleApiUtil;

    public Map<Object, Object> readDataFromGoogleSheet() throws GeneralSecurityException, IOException {
        return googleApiUtil.getDataFromSheet();
    }

    public GoogleSheetResponseDTO createSheet(List<String> request) throws GeneralSecurityException, IOException {
        return googleApiUtil.createGoogleSheet(request);
    }

    public String updateSheet(GoogleSheetDTO request) throws GeneralSecurityException, IOException {
        return googleApiUtil.updateGoogleSheet(request);
    }
}
