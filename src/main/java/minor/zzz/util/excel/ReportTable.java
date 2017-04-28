package minor.zzz.util.excel;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by zhouzb on 2017/3/8.
 *
 * 非线程安全
 */
public class ReportTable {

    private static class Column {
        private String field;               // 列字段field
        private int columnIndex;            // 列索引
        private DataType dataType;          // 数据类型
    }

    public static class Title {
        private String mainTitle = StringUtils.EMPTY;
        private String subTitle = StringUtils.EMPTY;

        public String getMainTitle() {
            return mainTitle;
        }

        public String getSubTitle() {
            return subTitle;
        }
    }

    public static class Head {
        private String field = StringUtils.EMPTY;                   // 表头字段
        private String name = StringUtils.EMPTY;                    // 表头名称
        private int columnSpan = 1;                                 // 表头跨列数
        private int rowSpan = 1;                                    // 表头跨行数
        private int rowIndex = 0;                                   // 表头相对行索引
        private int columnIndex = 0;                                // 表头相对列索引
        private DataType dataType = DataType.STRING;                // 表头数据类型

        private List<Head> children = new ArrayList<Head>();        // 子表头

        public boolean isGroup() {
            return !CollectionUtils.isEmpty(children);
        }

        private int deepestLevel() {
            int depth = 1;
            if (isGroup()) {
                int max = 0;
                for (Head child : children) {
                    if (max < child.deepestLevel()) {
                        max = child.deepestLevel();
                    }
                }
                depth = depth + max;
            }

            return depth;
        }

        private void updateRowIndex(int start) {
            this.rowIndex = start + 1;
            if (this.isGroup()) {
                for (Head child : children) {
                    child.updateRowIndex(this.rowIndex);
                }
            }
        }

        private void updateSelf(Head child) {

            int childRowSpan = child.deepestLevel();
            int childColumnSpan = child.columnSpan;

            if (childRowSpan < this.rowSpan) {

            } else {
                this.rowSpan = childRowSpan + 1;
            }

            if (CollectionUtils.isEmpty(this.children)) {
                this.columnSpan = childColumnSpan;
            } else {
                this.columnSpan += childColumnSpan;
            }
        }

        private void updateChild(Head child) {
            Head closetSiblings = null;
            if (CollectionUtils.isEmpty(this.children)) {
                child.columnIndex = 0;
            } else {
                closetSiblings = this.children.get(this.children.size() - 1);
                child.columnIndex = closetSiblings.columnIndex + closetSiblings.columnSpan;
            }

            child.updateRowIndex(this.rowIndex);
        }

        public void addChild(Head child) {
            updateSelf(child);

            updateChild(child);

            this.children.add(child);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;

            Head head = (Head) o;

            if (!field.equals(head.field)) return false;
            return name.equals(head.name);
        }

        @Override
        public int hashCode() {
            int result = field.hashCode();
            result = 31 * result + name.hashCode();
            return result;
        }
    }


    public static class Footer {

    }

    public enum DataType {
        STRING, INTEGER, DOUBLE, DATE;
    }

    public static class Data extends HashMap<String, Object> {

    }

    private Title title;
    private List<Head> head = new ArrayList<>();
    private Footer footer;
    private List<Data> data = new ArrayList<>();
    private List<String> groupField = new ArrayList<>();
    private boolean dataChanged = false;
    private boolean groupFieldChanged = false;

    public Title getTitle() {
        return this.title;
    }

    public List<Head> getHead() {
        return head;
    }

    public Footer getFooter() {
        return footer;
    }

    public List<Data> getData() {
        if (CollectionUtils.isEmpty(this.data)) {
            return new ArrayList<Data>();
        }

        return this.data;
    }

    public List<String> getGroupField() {
        if (CollectionUtils.isEmpty(this.groupField)) {
            return new ArrayList<String>();
        }

        return this.groupField;
    }

    // 数据分组列
    public void setGroupField(List<String> groupField) {
        if (this.groupField != groupField) {
            groupFieldChanged = true;

            if (CollectionUtils.isEmpty(groupField)) {
                groupField = new ArrayList<String>();
            }

            this.groupField = groupField;

            sortedRawData();
        }

    }

    // 数据
    public void setData(List<Data> data) {

        if (this.data != data) {
            dataChanged = true;

            if (CollectionUtils.isEmpty(data)) {
                data = new ArrayList<Data>();
            }

            this.data = data;

            sortedRawData();
        }
    }

    public void setHead(List<Head> head) {
        if (CollectionUtils.isEmpty(head)) {
            head = new ArrayList<Head>();
        }

        this.head = head;
    }

    // 排序
    private void sortedRawData() {
        if (CollectionUtils.isEmpty(this.data) || CollectionUtils.isEmpty(groupField)) {
            return;
        }

        if (this.dataChanged || this.groupFieldChanged) {

            Collections.sort(this.data, new Comparator<Data>() {
                @Override
                public int compare(Data data1, Data data2) {
                    Object v1, v2;
                    String field;
                    int result = 0;
                    Comparable cv1, cv2;
                    for (int i = 0; i < groupField.size(); i ++) {
                        field = groupField.get(i);
                        v1 = data1.get(field);
                        v2 = data2.get(field);

                        cv1 = (Comparable)v1;
                        cv2 = (Comparable)v2;

                        if (cv1 == null && cv2 == null) {
                            result = 0;
                        } else if (cv1 == null && cv2 != null) {
                            result = -1;
                        } else if (cv1 != null && cv2 == null) {
                            result = 1;
                        } else {
                            result = cv1.compareTo(cv2);
                        }

                        if (result != 0) {
                            break;
                        }
                    }

                    return result;
                }
            });

            dataChanged = false;
            groupFieldChanged = false;
        }
    }

    // 导出excel
    public void exportExcel() {
        ExcelGenerator.setReportTable(this);
        ExcelGenerator.generate();
    }

    private static class ExcelGenerator {
        private static ReportTable table;

        private static void setReportTable(ReportTable reportTable) {
            table = reportTable;
        }

        private static int nextRowNum(HSSFSheet sheet) {
            int rowNum = 0;

            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();

            if (firstRowNum == lastRowNum) {
                rowNum = 0;
            } else {
                rowNum = lastRowNum + 1;
            }

            return rowNum;
        }

        private static int deepestNumOfHead() {
            int depth = -1;

            for (Head head : table.getHead()) {
                if (depth < head.rowSpan) {
                    depth = head.rowSpan;
                }
            }

            return depth;
        }

        private static void mergeCell(HSSFSheet sheet, int startRow, int endRow, int startColumn, int endColumn) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(startRow, endRow, startColumn, endColumn);
            sheet.addMergedRegion(cellRangeAddress);
        }

        private static void generateTitle(HSSFSheet sheet) {
            HSSFRow row = sheet.createRow(0);
            int colNum = 0;
            HSSFCell cell = null;
            if (StringUtils.isNotBlank(table.getTitle().getMainTitle())) {
                cell = row.createCell(colNum++);
                cell.setCellValue(table.getTitle().getMainTitle());
            }

            if (StringUtils.isNotBlank(table.getTitle().getSubTitle())) {
                cell = row.createCell(colNum++);
                cell.setCellValue(table.getTitle().getSubTitle());
            }
        }

        private static void generateHead(HSSFSheet sheet,
                                         HSSFCellStyle cellStyle,
                                         List<Head> head,
                                         Map<Integer, HSSFRow> rowMap,
                                         Map<String, Column> field2column,
                                         int rowStart, int columnStart) {
            int colSpan, rowIndex;
            HSSFRow row;
            HSSFCell cell;
            Column column;
            for (Head h : head) {
                colSpan = h.columnSpan;

                rowIndex = h.rowIndex;

                row = rowMap.get(rowIndex);

                if (row == null) {
                    row = sheet.createRow(rowStart + rowIndex);
                    rowMap.put(rowIndex, row);
                }

                cell = row.createCell(columnStart);
                cell.setCellValue(h.name);
                cell.setCellStyle(cellStyle);

                if (h.isGroup()) {
                    generateHead(sheet, cellStyle, h.children, rowMap, field2column, rowStart, columnStart);

                    mergeCell(sheet, rowStart + h.rowIndex, rowStart + h.rowIndex, columnStart, columnStart + colSpan - 1);
                } else {
                    column = new Column();
                    column.field = h.field;
                    column.columnIndex = columnStart;
                    column.dataType = h.dataType;

                    field2column.put(h.field, column);

                    mergeCell(sheet, rowStart + h.rowIndex, rowStart + deepestNumOfHead() - 1, columnStart, columnStart + colSpan - 1);
                }

                columnStart += colSpan;
            }
        }

//        private static Map<String, Column> createHead(HSSFSheet sheet, HSSFCellStyle cellStyle) {
//            int headStart = nextRowNum(sheet);
//
//            Map<Integer, HSSFRow> rowMap = new HashMap<>();
//
//            Map<String, Column> field2column = new HashMap<>();
//
//            generateHead(sheet, cellStyle, table.getHead(), rowMap, field2column, headStart, 0);
//
//            return field2column;
//        }

        private static void generateBody(HSSFSheet sheet, Map<String, Column> field2column, List<Data> data, final List<String> groupField) {
            int dataRowStart = nextRowNum(sheet);

            int rowNum = dataRowStart;

            Map<String, Object> groupFieldValue = new HashMap<>();
            Map<String, Integer> groupFieldRow = new HashMap<>();

            String field;
            Object value, groupValue;
            Column column;
            HSSFRow row;
            HSSFCell cell;

            for (Data _d : data) {

                row = sheet.createRow(rowNum);

                for (Map.Entry<String, Object> entry : _d.entrySet()) {
                    field = entry.getKey();
                    value = entry.getValue();

                    column = field2column.get(field);

                    cell = row.createCell(column.columnIndex);
                    cell.setCellValue(String.valueOf(value));

                    if (table.getGroupField().contains(field)) {
                        if (!groupFieldValue.containsKey(field)) {
                            groupFieldValue.put(field, value);
                            groupFieldRow.put(field, rowNum);
                        } else {
                            groupValue = groupFieldValue.get(field);

                            if (groupValue != null) {
                                if (!groupValue.equals(value)) {
                                    mergeCell(sheet, groupFieldRow.get(field), rowNum - 1, column.columnIndex, column.columnIndex);

                                    groupFieldValue.put(field, value);
                                    groupFieldRow.put(field, rowNum);
                                }
                            } else {
                                groupFieldValue.put(field, value);
                                groupFieldRow.put(field, rowNum);
                            }
                        }
                    }
                }

                rowNum ++;
            }

            // 收尾
            for (Map.Entry<String, Integer> entry : groupFieldRow.entrySet()) {
                field = entry.getKey();

                mergeCell(sheet, entry.getValue(), rowNum - 1, field2column.get(field).columnIndex, field2column.get(field).columnIndex);
            }
        }

        private static HSSFCellStyle cellStyle4Head(HSSFWorkbook excel) {
            HSSFCellStyle cellStyle = excel.createCellStyle();

            cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

            HSSFFont font = excel.createFont();
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示

            cellStyle.setFont(font);

            return cellStyle;
        }

        private static void generate() {
            HSSFWorkbook excel = new HSSFWorkbook();

            HSSFSheet sheet = excel.createSheet("excel");

            HSSFCellStyle cellStyle = cellStyle4Head(excel);

            Map<Integer, HSSFRow> rowMap = new HashMap<>();

            Map<String, Column> field2column = new HashMap<>();

            // 生成标题
//        createTitle(sheet);

            // 生成列头
            int headStart = nextRowNum(sheet);

            headStart = 4;

            generateHead(sheet, cellStyle, table.getHead(), rowMap, field2column, headStart, 0);

            // 填充数据
            generateBody(sheet, field2column, table.getData(), table.getGroupField());

            // 生成
            FileOutputStream fout = null;
            try{
                fout = new FileOutputStream("D:/students.xls");
                excel.write(fout);
                fout.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    private class HtmlTableGenerator {

        private String generate() {


            return null;
        }
    }

    public static void main(String[] args) {
        Head h = new Head();
        h.field = "h";
        h.name = "h";

        Head sh = new Head();
        sh.field = "sh";
        sh.name = "sh";

        Head ssh = new Head();
        ssh.field = "ssh";
        ssh.name = "ssh";

        Head ssh2 = new Head();
        ssh2.field = "ssh2";
        ssh2.name = "ssh2";

        sh.addChild(ssh);
        sh.addChild(ssh2);

        Head sh2 = new Head();
        sh2.field = "sh2";
        sh2.name = "sh2";
        Head sh3= new Head();
        sh3.field = "sh3";
        sh3.name = "sh3";

        h.addChild(sh);
        h.addChild(sh2);
        h.addChild(sh3);

        Head j = new Head();
        j.field = "j";
        j.name = "j";

        Head sj = new Head();
        sj.field = "sj";
        sj.name = "sj";

        Head sj2 = new Head();
        sj2.field = "sj2";
        sj2.name = "sj2";


        j.addChild(sj);
        j.addChild(sj2);

        List<Head> heads = new ArrayList<>();

        heads.add(h);
        heads.add(j);

        List<String> groupField = new ArrayList<>();
        groupField.add("ssh");
        groupField.add("ssh2");

        ReportTable table = new ReportTable();

        table.setHead(heads);
        table.setGroupField(groupField);

        List<Data> datas = new ArrayList<Data>();

        Data data;

        for (int i = 0; i < 5; i ++) {
            data = new Data();

            if (i == 0 || i == 4) {
                data.put("ssh", 0);
            } else {
                data.put("ssh", 1);
            }

            if (i == 1 || i == 3) {
                data.put("ssh2", 9);
            } else {
                data.put("ssh2", i + 10);
            }

            data.put("sj", "sj" + i);
            data.put("sj2", i);
            data.put("sh2", i);
            data.put("sh3", i);

            datas.add(data);

        }

        table.setData(datas);
        table.exportExcel();
    }
}
