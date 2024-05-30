package com.nisum.google.sheets.util;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.nisum.google.sheets.dto.GoogleSheetDTO;
import com.nisum.google.sheets.dto.GoogleSheetResponseDTO;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.LongStream;

@Slf4j
@Component
public class GoogleApiUtil {

    public static final String DATE_FORMAT = "ddMMyyyy";
    @Value("${google.spreadsheet.id}")
    private String spreadsheetId;

    @Value("${google.spreadsheet.headers}")
    private String spreadsheetHeaders;

    private static final String APPLICATION_NAME = "Google Sheets API Integration";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens/path";

    /**
     * Global instance of the scopes required by this quickstart. If modifying these
     * scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS, SheetsScopes.DRIVE);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    /**
     * Creates an authorized Credential object.
     *
     * @param httpTransport The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport httpTransport) throws IOException {
        // Load client secrets.
        InputStream in = GoogleApiUtil.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(httpTransport, JSON_FACTORY,
                clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(
                        new java.io.File(System.getProperty("user.home"), TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline").build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public Map<Object, Object> getDataFromSheet() throws GeneralSecurityException, IOException {
        List<List<Object>> values = getDataSheet();
        Map<Object, Object> data = new HashMap<>();
        if (values == null || values.isEmpty()) {
            log.error("No data found.");
        } else {
            for (List row : values) {
                data.put(row.get(0), row.get(1));
            }
        }
        return data;
    }

    private List<List<Object>> getDataSheet() throws GeneralSecurityException, IOException {
        LocalDate today = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_FORMAT);
        String formattedDate = today.format(formatter);
        final String range = "EMP-" + formattedDate + "!A2:D";
        Sheets service = getSheetService();
        ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
        return response.getValues();
    }

    private Sheets getSheetService() throws GeneralSecurityException, IOException {
        final NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
        return new Sheets.Builder(httpTransport, JSON_FACTORY, getCredentials(httpTransport))
                .setApplicationName(APPLICATION_NAME).build();
    }

    private Drive getDriveService() throws GeneralSecurityException, IOException {
        final NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
        return new Drive.Builder(httpTransport, JSON_FACTORY, getCredentials(httpTransport))
                .setApplicationName(APPLICATION_NAME).build();
    }

    public GoogleSheetResponseDTO createGoogleSheet(List<String> emails) throws GeneralSecurityException, IOException {
        LocalDate start = LocalDate.of(LocalDate.now().getYear(), LocalDate.now().getMonth(), LocalDate.now().getDayOfMonth());
        LocalDate end = LocalDate.of(LocalDate.now().getYear(), 12, 31);
        long days = ChronoUnit.DAYS.between(start, end) + 1;
        List<Sheet> sheets = new ArrayList<>();
        LongStream.range(0, days).forEach(day -> createSheets(day, start, sheets));
        SpreadsheetProperties spreadsheetProperties = new SpreadsheetProperties();
        spreadsheetProperties.setTitle("Employee-Security-" + LocalDate.now().getYear());
        Spreadsheet spreadsheet = new Spreadsheet().setProperties(spreadsheetProperties).setSheets(sheets);
        Sheets service = getSheetService();
        Spreadsheet createdResponse = service.spreadsheets().create(spreadsheet).execute();

        LocalDate today = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_FORMAT);
        String formattedDate = today.format(formatter);
        final String range = "EMP-" + formattedDate + "!A1";

        List<Object> headers = Arrays.asList(spreadsheetHeaders.split(","));
        ValueRange valueRange = new ValueRange().setValues(List.of(headers));
        service.spreadsheets().values().update(createdResponse.getSpreadsheetId(), range, valueRange).setValueInputOption("RAW").execute();

        final var googleSheetResponseDTO = new GoogleSheetResponseDTO();
        googleSheetResponseDTO.setSpreadSheetId(createdResponse.getSpreadsheetId());
        googleSheetResponseDTO.setSpreadSheetUrl(createdResponse.getSpreadsheetUrl());

        Drive driveService = getDriveService();
        emails.forEach(emailAddress -> providePermissions(emailAddress, driveService, createdResponse));
        return googleSheetResponseDTO;
    }

    private static void providePermissions(String emailAddress, Drive driveService, Spreadsheet createdResponse) {
        Permission permission = new Permission().setType("user").setRole("writer").setEmailAddress(emailAddress);
        try {
            driveService.permissions().create(createdResponse.getSpreadsheetId(), permission).setSendNotificationEmail(true).setEmailMessage("Google Sheet Permission");
        } catch (IOException e) {
            log.error("Error while creating permission for the user: {}", emailAddress, e);
        }
    }

    private static void createSheets(long i, LocalDate start, List<Sheet> sheets) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_FORMAT);
        String formattedDate = start.plusDays(i).format(formatter);
        SheetProperties sheetProperties = new SheetProperties();
        sheetProperties.setTitle("EMP-" + formattedDate);
        Sheet sheet = new Sheet().setProperties(sheetProperties);
        sheets.add(sheet);
    }


    public String updateGoogleSheet(GoogleSheetDTO request) throws GeneralSecurityException, IOException {
        List<List<Object>> data = getDataSheet();
        int size = ((data == null || data.isEmpty()) ? 2 : data.size()) + 2;

        LocalDate today = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_FORMAT);
        String formattedDate = today.format(formatter);
        final String range = "EMP-" + formattedDate + "!A" + size;

        String loginTime = LocalTime.now(ZoneId.of("Asia/Kolkata")).toString();
        request.getData().forEach(row -> row.add(loginTime));

        Sheets service = getSheetService();
        ValueRange valueRange = new ValueRange().setValues(request.getData());
        service.spreadsheets().values().update(spreadsheetId, range, valueRange).setValueInputOption("RAW").execute();

        return "Success | " + loginTime;
    }
}
