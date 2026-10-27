import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Cleaned-up version of the Good Times attendance final print program.
 * Produces the same grouped, sorted, page-padded report as the original
 * while eliminating massive duplication.
 */
public class FinalPrintClean {

    private static final String BASE_PATH = "C:/Users/Good Times/Desktop/ProgramAttendance/";
    private static final String[] LOCATION_FILES = {
        "HWY18.xls", "TT.xls", "Brookfield.xls", "Delafield.xls"
    };
    private static final String MASTER_FILE = "Read.xls";
    private static final int SENTINEL = 999999;
    private static final int PAGE_LINES = 54;

    // Exact counselor strings as they appear in the spreadsheet
    private static final Map<String, List<String>> GROUPS = new LinkedHashMap<>();
    static {
        GROUPS.put("|   Forever Fun Girls   |", Arrays.asList("LAUREN", "LAURA", "SARA", "ANNA"));
        GROUPS.put("|   Forever Fun Boys    |", Arrays.asList("TOM", "SHAWN", "ROSS", "KARA", "SEAN"));
        GROUPS.put("|  Forever Young Girls  |", Arrays.asList("KIM", "CLAIRE", "AMY"));
        GROUPS.put("|  Forever Young Boys   |", Arrays.asList("PATRICK", "ALYSSA", "KIRSTEN", "LIZT"));
        GROUPS.put("| Junior Explorer Girls |", Arrays.asList("LIZ", "LEAH", "MICHELLE"));
        GROUPS.put("| Junior Explorer Boys  |", Arrays.asList("DAN W. ", "BEN", "DAN G. ", "MITCH"));
        GROUPS.put("|Seasoned Explorer Girls|", Arrays.asList("EMMA", "GRACE", "MELISSA", "MOLLY", "KELSEY", "ANDREA"));
        GROUPS.put("|Seasoned Explorer Boys |", Arrays.asList("JUSTIN", "MATT", "PETER", "RYAN"));
        GROUPS.put("|   Rising Star Girls   |", Arrays.asList("CHELSEA", "KAYLYN", "BECCA", "ALEAH"));
        GROUPS.put("|   Rising Star Boys    |", Arrays.asList("MIKE S.", "DAVID", "RICHARD", "BRIAN"));
        GROUPS.put("|  Shooting Star Girls  |", Arrays.asList("MEGAN", "CASSIE", "ALLI"));
        GROUPS.put("|  Shooting Star Boys   |", Arrays.asList("MIKE T.", "CHRIS"));
        GROUPS.put("|     Varsity Girls     |", Arrays.asList("SAM", "SARAH", "KATIE", "ANNE"));
        GROUPS.put("|     Varsity Boys      |", Arrays.asList("PJ", "NICK ", "NATHANIEL"));
        GROUPS.put("|   Senior Varsity      |", Arrays.asList("ASHLEY", "DUSTIN"));
    }

    public static void main(String[] args) throws IOException {
        // 1. Collect all attendance IDs from the location files
        List<Integer> attendanceIds = new ArrayList<>();
        for (String locFile : LOCATION_FILES) {
            attendanceIds.addAll(readIdsFromFile(BASE_PATH + locFile));
        }

        // 2. Build lookup: camper ID -> counselor name + camper name
        Map<Integer, String[]> master = loadMaster(BASE_PATH + MASTER_FILE);

        // 3. Bucket campers by counselor
        Map<String, List<String>> byCounselor = new HashMap<>();
        List<String> unknown = new ArrayList<>();
        for (int id : attendanceIds) {
            String[] info = master.get(id);
            if (info == null) continue;
            String counselor = info[0];
            String camper = info[1];
            byCounselor.computeIfAbsent(counselor, k -> new ArrayList<>()).add(camper);
        }

        // 4. Generate the report
        Date today = new Date();
        String formattedDate = new SimpleDateFormat("EEE, dd-MMM").format(today);
        String outFile = formattedDate + ".doc";

        try (FileOutputStream fos = new FileOutputStream(outFile);
             PrintStream out = new PrintStream(fos)) {

            for (Map.Entry<String, List<String>> group : GROUPS.entrySet()) {
                List<String> campers = collectCampers(byCounselor, group.getValue());
                Collections.sort(campers);
                boolean isSeniorVarsity = group.getKey().contains("Senior Varsity");
                printSection(out, group.getKey(), formattedDate, campers, false, isSeniorVarsity);
            }

            // Special Cath page (extra blank lines)
            List<String> cath = byCounselor.getOrDefault("Cath", Collections.emptyList());
            if (!cath.isEmpty()) {
                Collections.sort(cath);
                printSection(out, "|       Cath            |", formattedDate, cath, true, false);
            }

            // Unknown counselors
            for (Map.Entry<String, List<String>> e : byCounselor.entrySet()) {
                String key = e.getKey();
                if (!isKnownCounselor(key) && !"Cath".equals(key)) {
                    unknown.addAll(e.getValue());
                }
            }
            if (!unknown.isEmpty()) {
                Collections.sort(unknown);
                printSection(out, "|       ????            |", formattedDate, unknown, true, false);
            }
        }
    }

    private static List<Integer> readIdsFromFile(String path) throws IOException {
        List<Integer> ids = new ArrayList<>();
        try (HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(path))) {
            HSSFSheet sheet = wb.getSheet("Sheet1");
            HSSFRow row = sheet.getRow(0);
            if (row == null) return ids;
            for (int col = 0; ; col++) {
                HSSFCell cell = row.getCell((short) col);
                if (cell == null) break;
                int val = (int) cell.getNumericCellValue();
                if (val == SENTINEL) break;
                ids.add(val);
            }
        }
        return ids;
    }

    private static Map<Integer, String[]> loadMaster(String path) throws IOException {
        Map<Integer, String[]> map = new HashMap<>();
        try (HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(path))) {
            HSSFSheet sheet = wb.getSheet("Sheet1");
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                HSSFRow row = sheet.getRow(r);
                if (row == null) continue;
                HSSFCell idCell = row.getCell((short) 0);
                if (idCell == null) continue;
                int id = (int) idCell.getNumericCellValue();
                if (id == SENTINEL) break;
                HSSFCell counselorCell = row.getCell((short) 1);
                HSSFCell nameCell = row.getCell((short) 4);
                String counselor = counselorCell != null ? counselorCell.getStringCellValue() : "";
                String name = nameCell != null ? nameCell.getStringCellValue() : "";
                map.put(id, new String[]{counselor, name});
            }
        }
        return map;
    }

    private static List<String> collectCampers(Map<String, List<String>> byCounselor, List<String> counselors) {
        List<String> result = new ArrayList<>();
        for (String c : counselors) {
            List<String> list = byCounselor.get(c);
            if (list != null) result.addAll(list);
        }
        return result;
    }

    private static boolean isKnownCounselor(String name) {
        for (List<String> list : GROUPS.values()) {
            if (list.contains(name)) return true;
        }
        return false;
    }

    private static void printSection(PrintStream out, String header, String date,
                                     List<String> campers, boolean extraBlanks, boolean svTotal) {
        out.println("_________________________");
        out.println(header + "      " + date);
        out.println("_________________________");

        int blanksAfterHeader = extraBlanks ? 4 : 1;
        for (int i = 0; i < blanksAfterHeader; i++) out.println(" ");

        int pageCounter = 3 + blanksAfterHeader;
        for (String camper : campers) {
            out.println(camper);
            pageCounter++;
        }

        out.println(" ");
        out.println(" ");
        String totalLabel = svTotal ? "Total SV Campers: " : "Total Campers: ";
        out.println(totalLabel + campers.size());
        pageCounter += 3;

        // Pad to next page boundary (same logic as original)
        while (pageCounter >= PAGE_LINES) {
            pageCounter -= PAGE_LINES;
        }
        int pad = PAGE_LINES - pageCounter;
        for (int i = 0; i < pad; i++) out.println(" ");
    }
}
