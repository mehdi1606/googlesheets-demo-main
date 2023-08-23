package com.googlesheets.demo.config;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.docs.v1.Docs;
import com.google.api.services.docs.v1.DocsScopes;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
@Configuration
public class DriverConfig {



    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    @Value("${application.name}")
    private String applicationName;
    @Value("${credentials.file.path}")
    private String credentialsFilePath;
    @Value("${tokens.directory.path2}")
    private String tokensDirectoryPath2;



    public Credential authorize(List<String> scopes) throws IOException, GeneralSecurityException {
        InputStream in = GoogleAuthorizationConfig.class.getResourceAsStream(credentialsFilePath);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));


        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(tokensDirectoryPath2)))
                .setAccessType("offline")
                .build();

        LocalServerReceiver localServerReceiver = new LocalServerReceiver.Builder().setPort(8888).build();

        return new AuthorizationCodeInstalledApp(
                flow, localServerReceiver)
                .authorize("user");
    }

    public Drive getDriveService() throws IOException, GeneralSecurityException {
        List<String> scopes = Arrays.asList(
                DriveScopes.DRIVE_FILE,
                DriveScopes.DRIVE,
                DocsScopes.DOCUMENTS,
                SlidesScopes.PRESENTATIONS
        );

        Credential credential = authorize(scopes);

        return new Drive.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                JSON_FACTORY, credential)
                .setApplicationName(applicationName)
                .build();
    }
}
