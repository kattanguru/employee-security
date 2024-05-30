package com.nisum.google.sheets.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class GoogleSheetResponseDTO {

    private String spreadSheetId;
    private String spreadSheetUrl;

}
