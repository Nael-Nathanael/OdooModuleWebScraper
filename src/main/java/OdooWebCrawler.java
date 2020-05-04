import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;

public class OdooWebCrawler {

    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final Sheet sheet = workbook.createSheet("Master-0-1");
    private HashMap<String, String> hasilSatuApp = new HashMap<>();

    public static void main(String[] args) throws IOException {
        ArrayList<String> pageLinks = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            pageLinks.add(
                    "https://apps.odoo.com/apps/modules/category/Website/browse/page/" + i + "?price=Paid&series=12.0"
            );
        }
        new OdooWebCrawler().getPageLinks(pageLinks);
    }

    public void getPageLinks(ArrayList<String> urlList) throws IOException {
        int counter = 1;
        for (String URL : urlList) {

            String category = "";

            java.net.URL url = new URL(URL);
            String path = url.getPath();
            String[] pathArray = path.split("/");
            boolean found = false;

            for (String pather : pathArray) {
                if (found) {
                    category = pather;
                    break;
                } else {
                    if (pather.equalsIgnoreCase("category")) {
                        found = true;
                    }
                }
            }

            Document document = Jsoup.connect(URL).get();
            generateExcel();

            Elements allApp = document.select(".loempia_app_entry");
            for (Element anApp : allApp) {
                hasilSatuApp.clear();

                hasilSatuApp.put("Category", category);

                String title = anApp.select("a>div.loempia_app_entry_bottom>div>h5>b").text();
                hasilSatuApp.put("Nama Umum", title);

                StringBuilder author = new StringBuilder();
                Elements authorEl = anApp.select("a>div.loempia_app_entry_bottom>div.row>div.loempia_panel_author");
                for (Element e : authorEl) {
                    author.append(e.text());
                }
                hasilSatuApp.put("Author", author.toString());

                String fokusCharge = anApp.select("a>div.loempia_app_entry_bottom>div.row>div.loempia_panel_price>b").text();
                String charge = fokusCharge.equalsIgnoreCase("free") ? "Free" : "Paid";
                hasilSatuApp.put("Charge", charge);

                String price = anApp.select("a>div.loempia_app_entry_bottom>div.row>div.loempia_panel_price>b>span>span.oe_currency_value").text();
                if (price.equalsIgnoreCase("")) {
                    price = "0";
                }
                hasilSatuApp.put("Harga (USD)", price);

                String link = anApp.select("a").attr("abs:href");
                hasilSatuApp.put("Link", link);

                String summary = anApp.select(".loempia_panel_summary").text();
                hasilSatuApp.put("Fungsi / Link Fungsi", summary);

                hasilSatuApp = innerPage(link, hasilSatuApp);

                writeRow(counter++);
            }
        }

        cetak();
    }

    private void writeRow(int rowNumber) {
        System.out.println("Sedang mengisi baris ke " + rowNumber);
        Row row = sheet.createRow(rowNumber);

        CellStyle rowStyle = workbook.createCellStyle();
        rowStyle.setWrapText(true);
        row.setHeight((short) 1000);
        rowStyle.setVerticalAlignment(VerticalAlignment.TOP);

        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        rowStyle.setFont(font);

        Cell cell = row.createCell(0);
        cell.setCellValue(rowNumber);
        cell.setCellStyle(rowStyle);

        cell = row.createCell(1);
        cell.setCellValue("02 Mei 2020");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(2);
        cell.setCellValue("Fasilkom UI");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(3);
        cell.setCellValue("");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(4);
        cell.setCellValue("");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(5);
        cell.setCellValue(1706043670);
        cell.setCellStyle(rowStyle);

        cell = row.createCell(6);
        cell.setCellValue("Nathanael");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(7);
        cell.setCellValue(hasilSatuApp.get("Category"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(8);
        cell.setCellValue("");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(9);
        cell.setCellValue("");
        cell.setCellStyle(rowStyle);

        cell = row.createCell(10);
        cell.setCellValue(hasilSatuApp.get("Charge"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(11);
        cell.setCellValue(Float.parseFloat(hasilSatuApp.get("Harga (USD)")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(12);
        cell.setCellValue(hasilSatuApp.get("License"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(13);
        cell.setCellValue(hasilSatuApp.get("Nama Umum"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(14);
        cell.setCellValue(hasilSatuApp.get("Technical name"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(15);
        cell.setCellValue(hasilSatuApp.get("Author"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(16);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("8")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(17);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("9")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(18);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("10")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(19);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("11")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(20);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("12")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(21);
        cell.setCellValue(Integer.parseInt(hasilSatuApp.get("13")));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(22);
        cell.setCellValue(hasilSatuApp.get("Live Link/website"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(23);
        cell.setCellValue(hasilSatuApp.get("Link"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(24);
        cell.setCellValue(hasilSatuApp.get("Fungsi / Link Fungsi"));
        cell.setCellStyle(rowStyle);

        cell = row.createCell(25);
        cell.setCellValue(hasilSatuApp.get("Required Apps"));
        cell.setCellStyle(rowStyle);
    }

    public void generateExcel() {
        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 8);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("No");

        headerCell = header.createCell(1);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Tanggal");

        headerCell = header.createCell(2);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Sumber Data");

        headerCell = header.createCell(3);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Pilihan");

        headerCell = header.createCell(4);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Set Modul");

        headerCell = header.createCell(5);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("NPM");

        headerCell = header.createCell(6);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Nama");

        headerCell = header.createCell(7);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Category");

        headerCell = header.createCell(8);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Sub Category");

        headerCell = header.createCell(9);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Sub Sub Category");

        headerCell = header.createCell(10);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Charge");

        headerCell = header.createCell(11);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Harga (USD)");

        headerCell = header.createCell(12);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("License");

        headerCell = header.createCell(13);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Nama Umum");

        headerCell = header.createCell(14);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Technical name");

        headerCell = header.createCell(15);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Author");

        headerCell = header.createCell(16);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(8);

        headerCell = header.createCell(17);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(9);

        headerCell = header.createCell(18);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(10);

        headerCell = header.createCell(19);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(11);

        headerCell = header.createCell(20);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(12);

        headerCell = header.createCell(21);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue(13);

        headerCell = header.createCell(22);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Live Link/website");

        headerCell = header.createCell(23);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Link");

        headerCell = header.createCell(24);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Fungsi / Link Fungsi");

        headerCell = header.createCell(25);
        headerCell.setCellStyle(headerStyle);
        headerCell.setCellValue("Required Apps");
    }

    public void cetak() throws IOException {
        for (int i = 0; i < 26; i++) {
            sheet.autoSizeColumn(i, true);
            sheet.setColumnWidth(i, Math.min(sheet.getColumnWidth(i) + 600, 10000));
        }

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        System.out.println(path);
        String fileLoc = path.substring(0, path.length() - 1) + "hasil.xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLoc);
        this.workbook.write(outputStream);
        this.workbook.close();
    }

    public HashMap<String, String> innerPage(String url, HashMap<String, String> hasilSatuApp) throws IOException {
        Document document = Jsoup.connect(url).get();
        Elements fokus = document.select(".loempia_app_table");

        Elements head = fokus.select("thead");
        Element body = fokus.select("tbody").get(0);

        String technicalName = body.select("tr>td>code").text();
        hasilSatuApp.put("Technical name", technicalName);

        Elements bodyRow = body.select("tr>td");
        Elements headRow = head.select("tr>td");


        String licence = "";
        boolean found = false;
        for (Element td : bodyRow) {
            if (found) {
                licence = td.text();
                break;
            }
            if (td.select("b").text().equalsIgnoreCase("license")) {
                found = true;
            }
        }
        hasilSatuApp.put("License", licence);

        Elements badges = bodyRow.select(".badge.bg-beta.mr8");
        if (badges.text().contains("8")) {
            hasilSatuApp.put("8", "1");
        } else {
            hasilSatuApp.put("8", "0");
        }
        if (badges.text().contains("9")) {
            hasilSatuApp.put("9", "1");
        } else {
            hasilSatuApp.put("9", "0");
        }
        if (badges.text().contains("10")) {
            hasilSatuApp.put("10", "1");
        } else {
            hasilSatuApp.put("10", "0");
        }
        if (badges.text().contains("11")) {
            hasilSatuApp.put("11", "1");
        } else {
            hasilSatuApp.put("11", "0");
        }
        if (badges.text().contains("12")) {
            hasilSatuApp.put("12", "1");
        } else {
            hasilSatuApp.put("12", "0");
        }
        if (badges.text().contains("13")) {
            hasilSatuApp.put("13", "1");
        } else {
            hasilSatuApp.put("13", "0");
        }

        String website = "";
        found = false;
        for (Element td : bodyRow) {
            if (found) {
                website = td.select("a").text();
                break;
            }
            if (td.select("b").text().equalsIgnoreCase("website")) {
                found = true;
            }
        }
        hasilSatuApp.put("Live Link/website", website);

        StringBuilder requiredApps = new StringBuilder();
        found = false;
        for (Element td : headRow) {
            if (found) {
                Elements target = td.select("span");
                for (Element el : target) {
                    requiredApps.append(el.text()).append(", ");
                }
                requiredApps = new StringBuilder(requiredApps.substring(0, requiredApps.length() - 2));
                break;
            }
            if (td.select("b").text().equalsIgnoreCase("required apps")) {
                found = true;
            }
        }
        hasilSatuApp.put("Required Apps", requiredApps.toString());
        return hasilSatuApp;
    }
}