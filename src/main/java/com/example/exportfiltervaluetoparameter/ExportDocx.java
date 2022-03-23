package com.example.exportfiltervaluetoparameter;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.SneakyThrows;
import lombok.experimental.Accessors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class ExportDocx {
    @SneakyThrows
    @GetMapping("/export/docx")
    public void exportDocx() {
        String path = "C:\\Users\\BENH VIEN CONG NGHE\\Downloads\\test.docx";
        XWPFDocument document = new XWPFDocument(new FileInputStream(path));
        Map<String, Object> mapField = addDataObjectIntoParameter(getHocSinh());
        document.getParagraphs().forEach(it -> {
            exportDataIntoDocx(it, mapField);
        });
        FileOutputStream out = new FileOutputStream("C:\\Users\\BENH VIEN CONG NGHE\\Downloads\\test-doc.docx");
        document.write(out);
        out.close();
        document.close();
    }

    private void exportDataIntoDocx(XWPFParagraph paragraph, Map<String, Object> mapField) {
        String text = paragraph.getText();
        if (StringUtils.isNotBlank(text))
            for (Map.Entry<String, Object> entry : mapField.entrySet()) {
                if (text.contains(entry.getKey())) {
                    changeTextRow(paragraph, entry.getKey(), entry.getValue().toString(), true);
                }
            }
    }

    public void changeTextRow(XWPFParagraph paragraph, String key, String value, boolean boldText) {
        List<XWPFRun> xwpfRuns = paragraph.getRuns();
        for (int i = 0; i < xwpfRuns.size(); i++) {
            if (xwpfRuns.get(i).toString().trim().equals(key)) {
                paragraph.getRuns().get(i).setText(value, 0);
                paragraph.getRuns().get(i).setBold(boldText);
            }
        }
    }

    public static Map<String, Object> addDataObjectIntoParameter(Object obj) {
        Map<String, Object> map = new HashMap<>();
        for (Field field : obj.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            try {
                map.put(field.getName(), field.get(obj));
            } catch (Exception ignored) {
            }
        }
        return map;
    }

    public HocSinh getHocSinh() {
        return new HocSinh()
                .setIdCategory(2)
                .setNameStudent("Anh")
                .setClassName("Lop 12")
                .setCrushName("Hoang thi bich hong");
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    @Accessors(chain = true)
    public class HocSinh {
        private int idCategory;
        private String nameStudent;
        private String className;
        private String crushName;
    }
}
