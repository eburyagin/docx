/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.eburyagin.docx;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 *
 * @author Администратор
 */
public class DocxMaker {

    public static void main(String[] args) throws Docx4JException {
        
        WordprocessingMLPackage doc = WordprocessingMLPackage.load(new File("C:\\docs\\neocenter_consult_msk_001.docx"));
        Map<String, Object> params = new HashMap<>();
        params.put("agree_number", "333-222/111");
        params.put("agree_date_str", "22 января 2014 г.");
        params.put("client_fio", "Ивано Иван Иванович");

        List<Map<String, Object>> tt = new ArrayList<>();
        Map<String, Object> r = new HashMap<>();
        r.put("fio", "Брокеров Брокер Брокерович");
        r.put("polis_number", "P12345");
        r.put("polis_period", "c 10.10.10 по 12.12.12");
        r.put("sro_name", "Саморегулируемая организация 1");
        r.put("sro_address", "Адрес 1 2 3");
        tt.add(r);
        r = new HashMap<>();
        r.put("fio", "Сидоров Сидор Брокерович");
        r.put("polis_number", "P54657");
        r.put("polis_period", "c 11.11.11 по 12.12.12");
        r.put("sro_name", "Саморегулируемая организация 2");
        r.put("sro_address", "Адрес 7 7 7");
        tt.add(r);

        params.put("broker", tt);

        DocxTemplateProcessor template = new DocxTemplateProcessor(doc);
        template.process(params);
        doc.save(new File("C:\\docs\\neocenter_consult_msk_001.result.docx"));

    }

}
