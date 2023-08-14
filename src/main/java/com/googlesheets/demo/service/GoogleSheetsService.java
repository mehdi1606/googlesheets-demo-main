package com.googlesheets.demo.service;

import java.io.IOException;
import java.security.GeneralSecurityException;

public interface GoogleSheetsService {
    void getSpreadsheetValues(String spreadsheetId) throws IOException, GeneralSecurityException;

}
