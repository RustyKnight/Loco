/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.beamcommunications.loco;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.StringJoiner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author swhitehead
 */
public class Main {

    public static int DESCRIPTION_COLUMN = 0;
    public static int COMMENT_COLUMN = 1;
    public static int CATEGORY_COLUMN = 2;
    public static int TYPE_COLUMN = 3;
    public static int NAME_COLUMN = 4;

    public static void main(String[] args) throws IOException {
        new Main();
    }

    private Workbook workbook;
    private Sheet sheet;

    public Main() throws IOException {
        workbook = WorkbookFactory.create(new File("Loco.xlsx"));
        sheet = workbook.getSheetAt(0);
        int languagesStartColumn = 5;
        int languagesEndColumn = languagesStartColumn;

        Row row = sheet.getRow(0);
        while (row.getCell(languagesEndColumn) != null) {
            languagesEndColumn++;
        }
        System.out.println("languagesStartColumn = " + languagesStartColumn);
        System.out.println("languagesEndColumn = " + languagesEndColumn);
        for (int languageColumn = languagesStartColumn; languageColumn < languagesEndColumn; languageColumn++) {
            String name = row.getCell(languageColumn).getStringCellValue();
            List<Localisation> entries = mapLanguage(languageColumn);
            formatIos(name, entries);
            formatAndroid(name, entries);
        }
    }

    protected String getValue(int rowIndex, int column) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        Cell cell = row.getCell(column);
        if (cell == null) {
            return null;
        }
        return cell.getStringCellValue();
    }

    private String lastCategory;

    protected Key getKeyAt(int rowIndex) {
        String category = getValue(rowIndex, CATEGORY_COLUMN);
        String type = getValue(rowIndex, TYPE_COLUMN);
        String name = getValue(rowIndex, NAME_COLUMN);

        if (category == null && type == null && name == null) {
            return null;
        }

        if (category == null) {
            category = lastCategory;
        }
        lastCategory = category;

        return new Key(category, type, name);
    }

    protected String formatiOSKey(Key key) {
        StringJoiner sj = new StringJoiner(".");
        sj.add(toCamelCase(key.getCategory()));
        sj.add(toCamelCase(key.getType()));
        sj.add(toCamelCase(key.getName()));
        
        return sj.toString();
    }

    protected String formatAndroid(Key key) {
        StringJoiner sj = new StringJoiner(".");
        sj.add(key.getCategory().toLowerCase());
        sj.add(key.getType().toLowerCase());
        sj.add(key.getName().toLowerCase().replace(" ", "_"));
        
        return sj.toString();
    }

    public List<Localisation> mapLanguage(int languageColumn) {
        List<Localisation> entries = new ArrayList<>(25);
        int rowIndex = 1;
        boolean done = false;
        do {
            Key key = getKeyAt(rowIndex);
            if (key == null) {
                done = true;
                continue;
            }
            String iosKey = formatiOSKey(key);
            String androidKey = formatAndroid(key);
            if (iosKey == null && androidKey == null) {
                done = true;
                continue;
            }
            String description = getValue(rowIndex, DESCRIPTION_COLUMN);
            String comment = getValue(rowIndex, COMMENT_COLUMN);
            String value = getValue(rowIndex, languageColumn);
            
            System.out.println("value = " + value);

            Localisation loco = new Localisation(value, iosKey, androidKey, comment, description);

            entries.add(loco);
            rowIndex++;
        } while (!done);

        return entries;
    }

    protected String dateTimeStamp() {
        return LocalDateTime.now().format(DateTimeFormatter.ofLocalizedDateTime(FormatStyle.MEDIUM));
    }

    public void formatIos(String language, List<Localisation> entries) {
        Collections.sort(entries, new Comparator<Localisation>() {
            @Override
            public int compare(Localisation o1, Localisation o2) {
                return o1.getIosKey().compareTo(o2.getIosKey());
            }
        });
        try (BufferedWriter bw = new BufferedWriter(new FileWriter(new File(language + ".strings")))) {
            bw.write("/*");
            bw.newLine();
            bw.write("\tAutogenerated localisation for " + language);
            bw.newLine();
            bw.write("\tExported on " + dateTimeStamp());
            bw.newLine();
            bw.write("*/");
            bw.newLine();
            bw.newLine();
            for (Localisation loco : entries) {
                String description = loco.getDescription();
                String comment = loco.getComment();

                if (description != null) {
                    bw.write("// " + description);
                    bw.newLine();
                }

                String value = loco.getValue();
                String key = loco.getIosKey();
                bw.write("\"" + key + "\" = \"" + value + "\";");
                bw.newLine();
            }
        } catch (IOException exp) {
            exp.printStackTrace();
        }
    }

    public void formatAndroid(String language, List<Localisation> entries) {
        Collections.sort(entries, new Comparator<Localisation>() {
            @Override
            public int compare(Localisation o1, Localisation o2) {
                return o1.getIosKey().compareTo(o2.getIosKey());
            }
        });
        try (BufferedWriter bw = new BufferedWriter(new FileWriter(new File(language + ".xml")))) {
            bw.write("<!--");
            bw.newLine();
            bw.write("\tAutogenerated localisation for " + language);
            bw.newLine();
            bw.write("\tExported on " + dateTimeStamp());
            bw.newLine();
            bw.write("-->");
            bw.newLine();
            bw.write("<resources xmlns:xliff=\"urn:oasis:names:tc:xliff:document:1.2\">");
            bw.newLine();
            for (Localisation loco : entries) {
                String description = loco.getDescription();
                String comment = loco.getComment();

                if (description != null) {
                    bw.write("<!-- " + description + " -->");
                    bw.newLine();
                }

                String value = loco.getValue();
                String key = loco.getAndroidKey();
                bw.write("<string name=\"" + key + "\">");
                bw.newLine();
                bw.write("\t" + value);
                bw.newLine();
                bw.write("</string>");
                bw.newLine();
            }
            bw.write("</resources>");
        } catch (IOException exp) {
            exp.printStackTrace();
        }
    }

    public static String toCamelCase(final String init) {
        if (init == null) {
            return null;
        }

        final StringBuilder ret = new StringBuilder(init.length());

        for (final String word : init.split(" ")) {
            if (!word.isEmpty()) {
                if (ret.length() == 0) {
                    ret.append(word.substring(0, 1).toLowerCase());
                } else {
                    ret.append(word.substring(0, 1).toUpperCase());
                }
                ret.append(word.substring(1).toLowerCase());
            }
        }

        return ret.toString();
    }

    public class Localisation {

        private String value;
        private String iosKey;
        private String androidKey;
        private String comment;
        private String description;

        public Localisation(String key, String iosValue, String androidValue, String comment, String description) {
            this.value = key;
            this.iosKey = iosValue;
            this.androidKey = androidValue;
            this.comment = comment;
            this.description = description;
        }

        public String getValue() {
            return value;
        }

        public String getIosKey() {
            return iosKey;
        }

        public String getAndroidKey() {
            return androidKey;
        }

        public String getComment() {
            return comment;
        }

        public String getDescription() {
            return description;
        }

    }

    public class Key {

        private String category;
        private String type;
        private String name;

        public Key(String category, String type, String name) {
            this.category = category;
            this.type = type;
            this.name = name;
        }

        public String getCategory() {
            return category;
        }

        public String getType() {
            return type;
        }

        public String getName() {
            return name;
        }

        @Override
        public String toString() {
            return category + "; " + type + "; " + name;
        }

    }
}
