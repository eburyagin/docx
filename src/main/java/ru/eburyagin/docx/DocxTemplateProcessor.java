/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.eburyagin.docx;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;
import javax.xml.bind.JAXBElement;
import org.apache.log4j.Logger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTSdtCell;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.SdtElement;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.jvnet.jaxb2_commons.ppp.Child;

/**
 *
 * @author Администратор
 */
public class DocxTemplateProcessor {

    private final WordprocessingMLPackage doc;
    private Map<String, Object> params;
    private int idx = 0;
    private Tr templateTr = null, newTr = null;
    private List<Object> cells = null;
    private List<String> attrs = null;

    private static final Logger logger = Logger.getLogger(DocxTemplateProcessor.class);

    public DocxTemplateProcessor(WordprocessingMLPackage doc) {
        this.doc = doc;
    }

    public WordprocessingMLPackage process(Map<String, Object> params) {

        this.params = params;

        MainDocumentPart main = doc.getMainDocumentPart();

        processContent(main);

        return doc;
    }

    private void processContent(ContentAccessor content) {

        if (content instanceof Tbl) {

            idx = 0;

        } else if (content instanceof Tr) {

            templateTr = (Tr) content;

        }

        for (int i = 0; i < content.getContent().size(); i++) {

            Object o = content.getContent().get(i);

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof ContentAccessor) {

                processContent((ContentAccessor) o);

                // после выхода проверяем, возможно были обработаны строки таблицы (Tr)
                if (content instanceof Tbl) {

                    // если так, то нужно проверить idx > 0
                    if (idx > 0) {

                        // если верно, то в таблице были "поля".
                        // нужно удалить templateTr из таблицы 
                        content.getContent().remove(templateTr);

                        // имя коллекции
                        StringTokenizer token = new StringTokenizer(attrs.get(0), ".");
                        String collName = token.nextToken();

                        // сама коллекция из параметров
                        List<Map<String, Object>> vals = (List<Map<String, Object>>) params.get(collName);
                        
                        if (vals == null) return;
                        
                        for (Map<String, Object> val : vals) {

                            newTr = org.docx4j.XmlUtils.deepCopy(templateTr);

                            newTr.getContent().clear();

                            for (int j = 0; j < cells.size(); j++) {
                                token = new StringTokenizer(attrs.get(j), ".");
                                token.nextToken();
                                newTr.getContent().add(replaceCell(cells.get(j), val.get(token.nextToken())));
                            }

                            content.getContent().add(newTr);

                        }

                        cells.clear();
                        attrs.clear();

                        break;

                    }

                }

            } else {

                if (o instanceof SdtElement) {

                    if (o instanceof SdtRun) {

                        Object run = processRun((SdtRun) o);

                        if (run == null) {
                            content.getContent().remove(i);
                        } else {
                            content.getContent().set(i, run);
                        }

                    } else if (o instanceof CTSdtCell) {

                        // если тег простой, то работаем как с обычным полем
                        if (!((CTSdtCell) o).getSdtPr().getTag().getVal().contains(".")) {

                            // блокируем обработуку "табличного заполнения" в данной таблице
                            idx = -1;
                            
                            Object cell = processCell((CTSdtCell) o);

                            content.getContent().set(i, cell);

                        } else {

                            if (idx == 0) {

                                idx++;
                                cells = new ArrayList<>();
                                attrs = new ArrayList<>();

                            }

//                        Object cell = prepareCell((CTSdtCell) o);
//                        content.getContent().set(i, cell);
                            // названия полей 
                            attrs.add(((CTSdtCell) o).getSdtPr().getTag().getVal());

                            // здесь мы только подготавливаем замену полей
                            // само заполнение будет после цикла for
                            cells.add(prepareCell((CTSdtCell) o));

                        }

                    }

                }

                logger.info(o.getClass().getName());

            }

        }

    }

    private RPr getRunProps(SdtElement run) {

        RPr props = null;

        // поиск свойств, которые нужно будет назначить полю после присвоения значения.
        // в этих свойствах в том числе определен стиль и т.п.
        for (Object o : run.getSdtPr().getRPrOrAliasOrLock()) {

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof RPr) {

                props = (RPr) o;

            }

        }

        return props;

    }

    /**
     * Обработка обычного поля (Run)
     *
     * @param run
     * @return изменненный
     */
    private Object processRun(SdtRun run) {

        if (!params.containsKey(run.getSdtPr().getTag().getVal())) {

            return null;

        }
        
        Logger.getLogger(this.getClass()).debug("Processing a tag = " + run.getSdtPr().getTag().getVal());

        replaceRunProps(run.getSdtContent(), getRunProps(run));

        replaceRunText(run.getSdtContent(), params.get(run.getSdtPr().getTag().getVal()));

        return run.getSdtContent().getContent().get(0);

    }

    /**
     * Обработка обычного поля (Cell)
     *
     * @param cell
     * @return изменненный
     */
    private Object processCell(CTSdtCell cell) {

        if (!params.containsKey(cell.getSdtPr().getTag().getVal())) {

            return null;

        }

        replaceCellProps(cell.getSdtContent(), getRunProps(cell));

        replaceCellText(cell.getSdtContent(), params.get(cell.getSdtPr().getTag().getVal()));

        Object result = null;

        for (Object o : cell.getSdtContent().getContent()) {

            if (o instanceof JAXBElement) {

                if (((JAXBElement) o).getValue() instanceof Tc) {

                    result = ((JAXBElement) o).getValue();

                }

            }

        }

        return result;

    }

    private void replaceRunProps(ContentAccessor content, RPr props) {

        for (int i = 0; i < content.getContent().size(); i++) {

            Object o = content.getContent().get(i);

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof ContentAccessor) {

                if (o instanceof R) {

                    ((R) o).setRPr(props);

                } else {

                    replaceRunProps((ContentAccessor) o, props);

                }

            }

        }

    }

    private void replaceRunText(ContentAccessor content, Object val) {

        for (int i = 0; i < content.getContent().size(); i++) {

            Object o = content.getContent().get(i);

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof ContentAccessor) {

                replaceRunText((ContentAccessor) o, val);

            } else {

                if (o instanceof Text) {

                    ((Text) o).setValue((val == null ? "" : val.toString()));

                }

            }

        }

    }

    private Object prepareCell(CTSdtCell cell) {

        replaceCellProps(cell.getSdtContent(), getRunProps(cell));

        Object result = null;

        for (Object o : cell.getSdtContent().getContent()) {

            if (o instanceof JAXBElement) {

                if (((JAXBElement) o).getValue() instanceof Tc) {

                    result = ((JAXBElement) o).getValue();

                }

            }

        }

        return result;

    }

    private void replaceCellProps(ContentAccessor content, RPr props) {

        logger.info(content.getClass().getName() + " - replaceCellProps");

        for (int i = 0; i < content.getContent().size(); i++) {

            Object o = content.getContent().get(i);

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof ContentAccessor) {

                if (o instanceof R) {

                    ((R) o).setRPr(props);

                } else {

                    replaceCellProps((ContentAccessor) o, props);

                }

            }

        }
    }

    private Object replaceCell(Object cell, Object val) {

        Object result = null;

        if (cell instanceof Tc) {

            JAXBElement parent = (JAXBElement) ((Child) cell).getParent();

            JAXBElement copy = org.docx4j.XmlUtils.deepCopy(parent);

            Object o = copy.getValue();

            if (o instanceof ContentAccessor) {

                replaceCellText((ContentAccessor) o, val);

            }

            result = o;

        }

        return result;

        //return org.docx4j.XmlUtils.deepCopy();
    }

    private void replaceCellText(ContentAccessor content, Object val) {

        for (int i = 0; i < content.getContent().size(); i++) {

            Object o = content.getContent().get(i);

            if (o instanceof JAXBElement) {

                o = ((JAXBElement) o).getValue();

            }

            if (o instanceof ContentAccessor) {

                replaceCellText((ContentAccessor) o, val);

            } else {

                if (o instanceof Text) {

                    ((Text) o).setValue(val.toString());

                }

            }

        }
    }
}
