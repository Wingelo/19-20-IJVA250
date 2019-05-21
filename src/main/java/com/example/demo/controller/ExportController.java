package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;
    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id" + ";" + "Nom" + ";" + "Prenom" + ";" + "Date de Naissance");

        for (Client client : allClients) {
            writer.println(client.getId() + ";"
                    + client.getNom() + ";"
                    + client.getPrenom() + ";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellPrenom = headerRow.createCell(1);
        cellPrenom.setCellValue("Prénom");

        Cell cellNom = headerRow.createCell(2);
        cellNom.setCellValue("Nom");

        int iRow = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(client.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(client.getPrenom());

            Cell nom = row.createCell(2);
            nom.setCellValue(client.getNom());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void factureXLSXByClient(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        List<Facture> factures = factureService.findFacturesClient(clientId);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Facture");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellTotal = headerRow.createCell(1);
        cellTotal.setCellValue("Prix Total");

        int iRow = 1;
        for (Facture facture : factures) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(facture.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(facture.getTotal());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void factureXLSXAllClient(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        Workbook workbook = new XSSFWorkbook();
        for (Client client : allClients) {
            Sheet sheet = workbook.createSheet(client.getNom());
            Row headerRow = sheet.createRow(0);

            // Paramétre Nom, Prenom etc...
            Cell cellId = headerRow.createCell(0);
            cellId.setCellValue("Id");

            Cell cellPrenom = headerRow.createCell(1);
            cellPrenom.setCellValue("Prénom");

            Cell cellNom = headerRow.createCell(2);
            cellNom.setCellValue("Nom");

            Cell cellDateDeNaissance = headerRow.createCell(3);
            cellDateDeNaissance.setCellValue("Date de naissance");

            // Détails du client
            Row row = sheet.createRow(1);

            Cell id = row.createCell(0);
            id.setCellValue(client.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(client.getPrenom());

            Cell nom = row.createCell(2);
            nom.setCellValue(client.getNom());

            Cell dateNaissance = row.createCell(3);
            dateNaissance.setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));

            List<Facture> factures = factureService.findFacturesClient(client.getId());
            //Factures
            for (Facture facture : factures) {
                Sheet sheetFacture = workbook.createSheet("Facture "+ facture.getId().toString());
                Row headerRowFacture = sheetFacture.createRow(0);

                Cell cellNomProduit = headerRowFacture.createCell(0);
                cellNomProduit.setCellValue("Nom de produit");

                Cell cellPrixProduit = headerRowFacture.createCell(1);
                cellPrixProduit.setCellValue("Prix produit");

                Cell cellQuantiteProduit = headerRowFacture.createCell(2);
                cellQuantiteProduit.setCellValue("Quantité");

                Cell cellSousTotal = headerRowFacture.createCell(3);
                cellSousTotal.setCellValue("Sous total");

                int rowFacture = 1;
                for(LigneFacture lignefacture : facture.getLigneFactures()){
                    Row rowF = sheetFacture.createRow(rowFacture);

                    Cell cellProduit =rowF.createCell(0);
                    cellProduit.setCellValue(lignefacture.getArticle().getLibelle());

                    Cell cellPrix = rowF.createCell(1);
                    cellPrix.setCellValue(lignefacture.getArticle().getPrix());

                    Cell cellQuantite = rowF.createCell(2);
                    cellQuantite.setCellValue(lignefacture.getQuantite());

                    Cell cellSTotal = rowF.createCell(3);
                    cellSTotal.setCellValue(lignefacture.getSousTotal());

                    rowFacture = rowFacture + 1;
                }
                CellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.RED.getIndex());
                style.setFillPattern(FillPatternType.forInt(PatternFormatting.SOLID_FOREGROUND));
                Font font = workbook.createFont();
                font.setColor(IndexedColors.WHITE.getIndex());
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setFont(font);

                CellStyle styleResultat = workbook.createCellStyle();
                styleResultat.setFillForegroundColor(IndexedColors.RED.getIndex());
                styleResultat.setFillPattern(FillPatternType.forInt(PatternFormatting.SOLID_FOREGROUND));
                Font fontResultat = workbook.createFont();
                fontResultat.setColor(IndexedColors.WHITE.getIndex());
                styleResultat.setFont(fontResultat);

                Row rowTotal = sheetFacture.createRow(rowFacture);

                Cell cellTotal =rowTotal.createCell(0);
                cellTotal.setCellValue("Total");

                cellTotal.setCellStyle(style);
                sheetFacture.addMergedRegion(new CellRangeAddress(rowFacture,rowFacture,0,2));
                Cell cellPrixTotal = rowTotal.createCell(3);
                cellPrixTotal.setCellValue(facture.getTotal());
                cellPrixTotal.setCellStyle(styleResultat);
            }



        }
        workbook.write(response.getOutputStream());
        workbook.close();

    }
}
