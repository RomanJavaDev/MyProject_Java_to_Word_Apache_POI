


import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hpsf.UnicodeString;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;

import static org.apache.poi.util.StringUtil.getFromUnicodeLE;
import static org.apache.poi.util.StringUtil.mapMsCodepointString;

public class ApachePOITest {

    public static void main(String[] args) throws Exception {
        try {
            FileOutputStream fos = new FileOutputStream(new File("Reabilitation card.docx"));
            XWPFDocument myNewDoc = new XWPFDocument();





// Добавляется нумерация страниц
            CTP ctp = CTP.Factory.newInstance();
//this add page number incremental
            ctp.addNewR().addNewPgNum();

            XWPFParagraph codePara = new XWPFParagraph(ctp, myNewDoc);
            XWPFParagraph[] paragraphs = new XWPFParagraph[1];
            paragraphs[0] = codePara;
//position of number
            codePara.setAlignment(ParagraphAlignment.RIGHT);

            CTSectPr sectPr = myNewDoc.getDocument().getBody().addNewSectPr();


            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(myNewDoc, sectPr);
            headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);


            CTPageMar pageMar = sectPr.addNewPgMar();
            pageMar.setLeft(BigInteger.valueOf(1200L));
            pageMar.setTop(BigInteger.valueOf(423L));
            pageMar.setRight(BigInteger.valueOf(567L));
            pageMar.setBottom(BigInteger.valueOf(285L));


            XWPFParagraph parag1 = myNewDoc.createParagraph();
            XWPFRun run1 = parag1.createRun();
            XWPFRun run01 = parag1.createRun();

            run1.setText("Государственное автономное учреждение социального обслуживания");
            run01.setText("ООО «Рога и копыта»");
            parag1.setAlignment(ParagraphAlignment.CENTER);
            run1.setFontFamily("Times New Roman");
            run1.setFontSize(13);
            run1.setBold(true);
            run1.addBreak();
            parag1.setSpacingAfter(1000);
            run01.setFontFamily("Times New Roman");
            run01.setFontSize(13);
            run01.setBold(true);
            run01.addBreak();


            XWPFParagraph parag2 = myNewDoc.createParagraph();
            XWPFRun run2 = parag2.createRun();
            XWPFRun run02 = parag2.createRun();

            run2.setText("Карта получателя социальных услуг");
            parag2.setAlignment(ParagraphAlignment.CENTER);
            run2.setFontFamily("Times New Roman");
            run2.setFontSize(24);
            run2.setBold(true);
            run2.addBreak();
            run02.setText("в стационарной/полустационарной форме");
            run02.setFontFamily("Times New Roman");
            run02.setFontSize(24);
            run02.setBold(true);
            run02.addBreak();


            XWPFRun run21 = parag2.createRun();

            run21.setText("№___________");
            run21.setFontFamily("Times New Roman");
            run21.setFontSize(20);
            run21.addBreak();
            parag2.setSpacingAfter(1000);

            XWPFParagraph parag4 = myNewDoc.createParagraph();
            XWPFRun run4 = parag4.createRun();
            parag4.setAlignment(ParagraphAlignment.LEFT);

            run4.setText("Ф.И.О.  ______________________________________");
            run4.addBreak();

            run4.setText("              " + "______________________________________");

            run4.setFontFamily("Times New Roman");
            run4.setFontSize(20);
            run4.setBold(true);
            run4.addBreak();
            parag4.setSpacingAfter(500);

            XWPFParagraph parag5 = myNewDoc.createParagraph();
            XWPFRun run5 = parag5.createRun();
            XWPFRun run51 = parag5.createRun();
            XWPFRun run52 = parag5.createRun();
            XWPFRun run53 = parag5.createRun();
            parag5.setAlignment(ParagraphAlignment.LEFT);
            run5.setText("Дата поступления  ___________________________________");
            run5.addBreak();
            run5.addBreak();
            run51.setText("Дата выписки _______________________________________");
            run51.addBreak();
            run51.addBreak();
            run52.setText("Осмотр на педикулез и чесотку ________________________");
            run52.addBreak();
            run52.addBreak();
            run53.setText("При стационарном обслуживании - комната № ___________");

            run5.setFontFamily("Times New Roman");
            run5.setFontSize(18);
            run51.setFontFamily("Times New Roman");
            run51.setFontSize(18);
            run52.setFontFamily("Times New Roman");
            run52.setFontSize(18);
            run53.setFontFamily("Times New Roman");
            run53.setFontSize(18);
            parag5.setSpacingAfter(2000);



            XWPFParagraph parag6 = myNewDoc.createParagraph();
            XWPFRun run6 = parag6.createRun();

            run6.setText("С правилами внутреннего распорядка для получателей социальных услуг ООО «Рога и копыта», ознакомлен (-а)");
            parag6.setAlignment(ParagraphAlignment.LEFT);
            run6.setFontFamily("Times New Roman");
            run6.setFontSize(16);
            run6.addBreak();

            XWPFRun run61 = parag6.createRun();
            run61.setText("«__»____________2020 г. _____________    _____________________");
            run61.setFontFamily("Times New Roman");
            run61.setFontSize(16);
            run61.addBreak();
            XWPFRun run62 = parag6.createRun();
            run62.setText("                                                 подпись                     Фамилия. И. О.");
            run62.setFontFamily("Times New Roman");
            run62.setFontSize(16);


            XWPFParagraph parag7 = myNewDoc.createParagraph();
            XWPFRun run7 = parag7.createRun();

            run7.setText("I. Общие данные о клиенте");
            parag7.setAlignment(ParagraphAlignment.CENTER);
            parag7.setPageBreak(true);
            run7.setFontFamily("Times New Roman");
            run7.setFontSize(14);
            run7.setBold(true);
            run7.addBreak();

            XWPFParagraph parag8 = myNewDoc.createParagraph();

            parag8.setAlignment(ParagraphAlignment.LEFT);
            parag8.setSpacingBetween(1.15);
            parag8.setSpacingAfterLines(0);

            XWPFRun run8 = parag8.createRun();
            run8.setFontFamily("Times New Roman");
            run8.setFontSize(12);
            run8.setBold(true);
            run8.setText("1. Фамилия, имя, отчество ");
            XWPFRun run801 = parag8.createRun();
            run801.setFontFamily("Times New Roman");
            run801.setFontSize(12);
            run801.setText("______________________________________________________________");

            run801.addBreak();
            XWPFRun run802 = parag8.createRun();
            run802.setFontFamily("Times New Roman");
            run802.setFontSize(12);
            run802.setBold(true);
            run802.setText("2. Дата рождения ");
            XWPFRun run803 = parag8.createRun();
            run803.setFontFamily("Times New Roman");
            run803.setFontSize(12);
            run803.setText("___________________________ ");
            XWPFRun run804 = parag8.createRun();
            run804.setFontFamily("Times New Roman");
            run804.setFontSize(12);
            run804.setBold(true);
            run804.setText("3. Возраст ");
            XWPFRun run805 = parag8.createRun();
            run805.setFontFamily("Times New Roman");
            run805.setFontSize(12);
            run805.setText("(число полных лет) ________________");
            run805.addBreak();
            XWPFRun run806 = parag8.createRun();
            run806.setFontFamily("Times New Roman");
            run806.setFontSize(12);
            run806.setText("4. Пол");
            run806.setBold(true);
            XWPFRun run807 = parag8.createRun();
            run807.setFontFamily("Times New Roman");
            run807.setFontSize(12);
            run807.setText(" (мужской/женский)");
            run807.addBreak();
            XWPFRun run808 = parag8.createRun();
            run808.setFontFamily("Times New Roman");
            run808.setFontSize(12);
            run808.setText("5. Адрес места жительства:");
            run808.setBold(true);
            XWPFRun run809 = parag8.createRun();
            run809.setFontFamily("Times New Roman");
            run809.setFontSize(12);
            run809.setText(" ______________________________________________________________");
            run809.addBreak();
            XWPFRun run810 = parag8.createRun();
            run810.setFontFamily("Times New Roman");
            run810.setFontSize(12);
            run810.setText("_______________________________________________________________________________________");
            run810.addBreak();
            XWPFRun run811 = parag8.createRun();
            run811.setFontFamily("Times New Roman");
            run811.setFontSize(12);
            run811.setText("6. Место постоянной регистрации");
            run811.setBold(true);
            XWPFRun run812 = parag8.createRun();
            run812.setFontFamily("Times New Roman");
            run812.setFontSize(12);
            run812.setText(" (при совпадении реквизитов с местом жительства пункт не заполняется)" +
                    " ___________________________________________________________________________");
            run812.addBreak();
            XWPFRun run813 = parag8.createRun();
            run813.setFontFamily("Times New Roman");
            run813.setFontSize(12);
            run813.setText("_______________________________________________________________________________________");
            run813.addBreak();
            XWPFRun run814 = parag8.createRun();
            run814.setFontFamily("Times New Roman");
            run814.setFontSize(12);
            run814.setText("7. Контактные телефоны:");
            run814.setBold(true);
            XWPFRun run815 = parag8.createRun();
            run815.setFontFamily("Times New Roman");
            run815.setFontSize(12);
            run815.setText(" _______________________________________________________________");
            run815.addBreak();

            XWPFRun run816 = parag8.createRun();
            run816.setFontFamily("Times New Roman");
            run816.setFontSize(12);
            run816.setText("8. Паспорт гражданина:");
            run816.setBold(true);
            XWPFRun run817 = parag8.createRun();
            run817.setFontFamily("Times New Roman");
            run817.setFontSize(12);
            run817.setText(" серия _____________ номер ______________ Кем выдан ________________");
            run817.addBreak();
            XWPFRun run818 = parag8.createRun();
            run818.setFontFamily("Times New Roman");
            run818.setFontSize(12);
            run818.setText("_______________________________________________________________________________________");
            run818.addBreak();

            XWPFRun run819 = parag8.createRun();
            run819.setFontFamily("Times New Roman");
            run819.setFontSize(12);
            run819.setText("9. ИПР/ИПРА инвалида № ___________________ Дата разработки ИПР/ИПРА ________________");
            run819.setBold(true);
            run819.addBreak();
            XWPFRun run820 = parag8.createRun();
            run820.setFontFamily("Times New Roman");
            run820.setFontSize(12);
            run820.setText("Бюро социальной экспертизы № _____________ ");
            run820.addBreak();

            XWPFRun run821 = parag8.createRun();
            run821.setFontFamily("Times New Roman");
            run821.setFontSize(12);
            run821.setText("10. Получает социальные услуги в ООО «Рога и копыта»:");
            run821.setBold(true);
            XWPFRun run822 = parag8.createRun();
            run822.setFontFamily("Times New Roman");
            run822.setFontSize(12);
            run822.setText(" впервые/ повторно.");
            run822.addBreak();

            XWPFRun run823 = parag8.createRun();
            run823.setFontFamily("Times New Roman");
            run823.setFontSize(12);
            run823.setText("11. Динамика повторных курсов");
            run823.setBold(true);
            XWPFRun run824 = parag8.createRun();
            run824.setFontFamily("Times New Roman");
            run824.setFontSize(12);
            run824.setText(" социальной реабилитации в социально-реабилитационных отделениях государственных учреждений социального обслуживания" +
                            " ___________________________");
            run824.addBreak();
            XWPFRun run825 = parag8.createRun();
            run825.setFontFamily("Times New Roman");
            run825.setFontSize(12);
            run825.setText("_______________________________________________________________________________________");
            run825.addBreak();

            XWPFParagraph parag9 = myNewDoc.createParagraph();
            XWPFRun run9 = parag9.createRun();

            run9.setText("II. Социальная характеристика");
            parag9.setAlignment(ParagraphAlignment.CENTER);
            run9.setFontFamily("Times New Roman");
            run9.setFontSize(14);
            run9.setBold(true);
            run9.addBreak();

            XWPFParagraph parag10 = myNewDoc.createParagraph();
            parag10.setAlignment(ParagraphAlignment.LEFT);
            parag10.setSpacingBetween(1.15);
            parag10.setSpacingAfterLines(0);

            XWPFRun run10 = parag10.createRun();
            run10.setFontFamily("Times New Roman");
            run10.setFontSize(12);
            run10.setBold(true);
            run10.setText("12. Группа инвалидности");
            XWPFRun run1001 = parag10.createRun();
            run1001.setFontFamily("Times New Roman");
            run1001.setFontSize(12);
            run1001.setText(" (первая, вторая, третья) ___________________________________________");
            run1001.addBreak();

            XWPFRun run1002 = parag10.createRun();
            run1002.setFontFamily("Times New Roman");
            run1002.setFontSize(12);
            run1002.setBold(true);
            run1002.setText("13. Причина инвалидности:");
            XWPFRun run1003 = parag10.createRun();
            run1003.setFontFamily("Times New Roman");
            run1003.setFontSize(12);
            run1003.setText("  _____________________________________________________________");
            run1003.addBreak();

            XWPFRun run1004 = parag10.createRun();
            run1004.setFontFamily("Times New Roman");
            run1004.setFontSize(12);
            run1004.setBold(true);
            run1004.setText("14. Инвалидность установлена до даты:");
            XWPFRun run1005 = parag10.createRun();
            run1005.setFontFamily("Times New Roman");
            run1005.setFontSize(12);
            run1005.setText("  __________________________________________________");
            run1005.addBreak();

            XWPFRun run1006 = parag10.createRun();
            run1006.setFontFamily("Times New Roman");
            run1006.setFontSize(12);
            run1006.setBold(true);
            run1006.setText("15. Длительность инвалидности");
            XWPFRun run1007 = parag10.createRun();
            run1007.setFontFamily("Times New Roman");
            run1007.setFontSize(12);
            run1007.setText(" (количество лет)  __________________________________________");
            run1007.addBreak();

            XWPFRun run1008 = parag10.createRun();
            run1008.setFontFamily("Times New Roman");
            run1008.setFontSize(12);
            run1008.setBold(true);
            run1008.setText("16. Преимущественные нарушения:");
            XWPFRun run1009 = parag10.createRun();
            run1009.setFontFamily("Times New Roman");
            run1009.setFontSize(12);
            run1009.setText(" (нужное подчеркнуть)");
            run1009.addBreak();
            XWPFRun run1010 = parag10.createRun();
            run1010.setFontFamily("Times New Roman");
            run1010.setFontSize(12);
            run1010.setText("ПОДА (в/конечности, н/конечности), слуха, зрения, психики, прочее");
            run1010.addBreak();

            XWPFRun run1011 = parag10.createRun();
            run1011.setFontFamily("Times New Roman");
            run1011.setFontSize(12);
            run1011.setBold(true);
            run1011.setText("17. Используемые технические средства реабилитации:");
            XWPFRun run1012 = parag10.createRun();
            run1012.setFontFamily("Times New Roman");
            run1012.setFontSize(12);
            run1012.setText(" (нужное подчеркнуть)");
            run1012.addBreak();
            XWPFRun run1013 = parag10.createRun();
            run1013.setFontFamily("Times New Roman");
            run1013.setFontSize(12);
            run1013.setText("к/коляска, костыли, трость, не использует");
            run1013.addBreak();

            XWPFRun run1014 = parag10.createRun();
            run1014.setFontFamily("Times New Roman");
            run1014.setFontSize(12);
            run1014.setBold(true);
            run1014.setText("18. Занятость");
            XWPFRun run1015 = parag10.createRun();
            run1015.setFontFamily("Times New Roman");
            run1015.setFontSize(12);
            run1015.setText(" (место работы):  ____________________________________________________________");
            run1015.addBreak();

            XWPFRun run1016 = parag10.createRun();
            run1016.setFontFamily("Times New Roman");
            run1016.setFontSize(12);
            run1016.setBold(true);
            run1016.setText("19. Причина, по которой не работает");
            XWPFRun run1017 = parag10.createRun();
            run1017.setFontFamily("Times New Roman");
            run1017.setFontSize(12);
            run1017.setText(" (по состоянию здоровья, отсутствие желаемой работы, отсутствие работы вблизи дома, отсутствие специально созданных условий труда, не желает работать), другое: _______________________________________________________________________");
            run1017.addBreak();

            XWPFRun run1018 = parag10.createRun();
            run1018.setFontFamily("Times New Roman");
            run1018.setFontSize(12);
            run1018.setBold(true);
            run1018.setText("20. Уровень материального благосостояния:");
            XWPFRun run1019 = parag10.createRun();
            run1019.setFontFamily("Times New Roman");
            run1019.setFontSize(12);
            run1019.setText(" средний (удовлетворяет, не удовлетворяет), уровень прожиточного минимума, ниже прожиточного минимума");
            run1019.addBreak();

            XWPFRun run1020 = parag10.createRun();
            run1020.setFontFamily("Times New Roman");
            run1020.setFontSize(12);
            run1020.setBold(true);
            run1020.setText("21. Образование");
            XWPFRun run1021 = parag10.createRun();
            run1021.setFontFamily("Times New Roman");
            run1021.setFontSize(12);
            run1021.setText(" ________________________________________________________________________");
            run1021.addBreak();

            XWPFRun run1022 = parag10.createRun();
            run1022.setFontFamily("Times New Roman");
            run1022.setFontSize(12);
            run1022.setBold(true);
            run1022.setText("22. Профессиональная подготовка");
            XWPFRun run1023 = parag10.createRun();
            run1023.setFontFamily("Times New Roman");
            run1023.setFontSize(12);
            run1023.setText("  _______________________________________________________");
            run1023.addBreak();

            XWPFRun run1024 = parag10.createRun();
            run1024.setFontFamily("Times New Roman");
            run1024.setFontSize(12);
            run1024.setBold(true);
            run1024.setText("23. Семейное положение:");
            XWPFRun run1025 = parag10.createRun();
            run1025.setFontFamily("Times New Roman");
            run1025.setFontSize(12);
            run1025.setText(" одинокий, одиноко проживающий, семейный, наличие иждивенцев  да/нет, проживает с родственниками, помогающими в обслуживании, проживает с родственниками, не обеспечивающими помощь, другое ______________________________________________________");
            run1025.addBreak();

            XWPFParagraph parag11 = myNewDoc.createParagraph();
            XWPFRun run11 = parag11.createRun();


            parag11.setAlignment(ParagraphAlignment.CENTER);
            parag11.setPageBreak(true);
            run11.setText("Проведение диагностики функционального состояния");
            run11.addBreak();
            run11.setText("(первичный осмотр)");
            run11.setFontFamily("Times New Roman");
            run11.setFontSize(16);
            run11.setBold(true);
            run11.addBreak();

            XWPFParagraph parag12 = myNewDoc.createParagraph();

            parag12.setAlignment(ParagraphAlignment.LEFT);
            parag12.setSpacingBetween(1.35);
            parag12.setSpacingAfterLines(0);


            XWPFRun run1201 = parag12.createRun();
            run1201.setFontFamily("Times New Roman");
            run1201.setFontSize(12);
            run1201.setBold(true);
            run1201.setText("24. Дата");
            XWPFRun run1202 = parag12.createRun();
            run1202.setFontFamily("Times New Roman");
            run1202.setFontSize(12);
            run1202.setText(" ____________________");
            run1202.addBreak();

            XWPFRun run1203 = parag12.createRun();
            run1203.setFontFamily("Times New Roman");
            run1203.setFontSize(12);
            run1203.setBold(true);
            run1203.setText("25. Основной диагноз:");
            run1203.addBreak();
            XWPFRun run1204 = parag12.createRun();
            run1204.setFontFamily("Times New Roman");
            run1204.setFontSize(12);
            run1204.setText(" ______________________________________________________________________________________");
            run1204.addBreak();
            run1204.setText(" ______________________________________________________________________________________");
            run1204.addBreak();

            XWPFRun run1205 = parag12.createRun();
            run1205.setFontFamily("Times New Roman");
            run1205.setFontSize(12);
            run1205.setBold(true);
            run1205.setText("26. Сопутствующий диагноз:");
            run1205.addBreak();
            XWPFRun run1206 = parag12.createRun();
            run1206.setFontFamily("Times New Roman");
            run1206.setFontSize(12);
            run1206.setText(" ______________________________________________________________________________________");
            run1206.addBreak();
            run1206.setText(" ______________________________________________________________________________________");
            run1206.addBreak();
            run1206.setText(" ______________________________________________________________________________________");
            run1206.addBreak();
            run1206.setText(" ______________________________________________________________________________________");
            run1206.addBreak();

            XWPFRun run1207 = parag12.createRun();
            run1207.setFontFamily("Times New Roman");
            run1207.setFontSize(12);
            run1207.setBold(true);
            run1207.setText("27. Жалобы:");
            run1207.addBreak();
            XWPFRun run1208 = parag12.createRun();
            run1208.setFontFamily("Times New Roman");
            run1208.setFontSize(12);
            run1208.setText(" ______________________________________________________________________________________");
            run1208.addBreak();
            run1208.setText(" ______________________________________________________________________________________");
            run1208.addBreak();
            run1208.setText(" ______________________________________________________________________________________");
            run1208.addBreak();
            run1208.setText(" ______________________________________________________________________________________");
            run1208.addBreak();

            XWPFRun run1209 = parag12.createRun();
            run1209.setFontFamily("Times New Roman");
            run1209.setFontSize(12);
            run1209.setBold(true);
            run1209.setText("28. История основного заболевания:");
            run1209.addBreak();
            XWPFRun run1210 = parag12.createRun();
            run1210.setFontFamily("Times New Roman");
            run1210.setFontSize(12);
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();
            run1210.setText(" ______________________________________________________________________________________");
            run1210.addBreak();

            XWPFRun run1211 = parag12.createRun();
            run1211.setFontFamily("Times New Roman");
            run1211.setFontSize(12);
            run1211.setBold(true);
            run1211.setText("29. История жизни, перенесенные заболевания:");
            run1211.addBreak();
            XWPFRun run1212 = parag12.createRun();
            run1212.setFontFamily("Times New Roman");
            run1212.setFontSize(12);
            run1212.setText(" ______________________________________________________________________________________");
            run1212.addBreak();
            run1212.setText(" ____________________________________________________________________________ вирусный гепатит ______________, туберкулез ________________, вен. заболевания _______________________");
            run1212.addBreak();

            XWPFRun run1213 = parag12.createRun();
            run1213.setFontFamily("Times New Roman");
            run1213.setFontSize(12);
            run1213.setBold(true);
            run1213.setText("30. Аллергический анамнез");

            XWPFRun run1214 = parag12.createRun();
            run1214.setFontFamily("Times New Roman");
            run1214.setFontSize(12);
            run1214.setText("  _____________________________________________________________");
            run1214.addBreak();

            XWPFRun run1215 = parag12.createRun();
            run1215.setFontFamily("Times New Roman");
            run1215.setFontSize(12);
            run1215.setBold(true);
            run1215.setText("31. Лекарственная непереносимость");

            XWPFRun run1216 = parag12.createRun();
            run1216.setFontFamily("Times New Roman");
            run1216.setFontSize(12);
            run1216.setText("  _____________________________________________________");
            run1216.addBreak();

            XWPFRun run1217 = parag12.createRun();
            run1217.setFontFamily("Times New Roman");
            run1217.setFontSize(12);
            run1217.setBold(true);
            run1217.setText("32. Гемотрансфузии");

            XWPFRun run1218 = parag12.createRun();
            run1218.setFontFamily("Times New Roman");
            run1218.setFontSize(12);
            run1218.setText(" __________________________");

            XWPFRun run1219 = parag12.createRun();
            run1219.setFontFamily("Times New Roman");
            run1219.setFontSize(12);
            run1219.setBold(true);
            run1219.setText(" Флюорография");

            XWPFRun run1220 = parag12.createRun();
            run1220.setFontFamily("Times New Roman");
            run1220.setFontSize(12);
            run1220.setText(" ___________________________");
            run1220.addBreak();

            XWPFRun run1221 = parag12.createRun();
            run1221.setFontFamily("Times New Roman");
            run1221.setFontSize(12);
            run1221.setBold(true);
            run1221.setText("33. Вредные привычки");

            XWPFRun run1222 = parag12.createRun();
            run1222.setFontFamily("Times New Roman");
            run1222.setFontSize(12);
            run1222.setText(" (подчеркнуть): курит, не курит, алкогольные напитки — не употребляет, употребляет умеренно, злоупотребляет; токсикомания, наркомания: да, нет, др.");
            run1222.addBreak();
            run1222.setText(" ______________________________________________________________________________________");
            run1222.addBreak();

            XWPFParagraph parag13 = myNewDoc.createParagraph();

            parag13.setAlignment(ParagraphAlignment.LEFT);
            parag13.setPageBreak(true);
            parag13.setSpacingBetween(1.35);
            parag13.setSpacingAfterLines(0);

            XWPFRun run1301 = parag13.createRun();
            run1301.setFontFamily("Times New Roman");
            run1301.setFontSize(12);
            run1301.setBold(true);
            run1301.setText("34. Эпиданамнез:");
            XWPFRun run1302 = parag13.createRun();
            run1302.setFontFamily("Times New Roman");
            run1302.setFontSize(12);
            run1302.setText(" контакт с инфекционными больными отрицает,");
            run1302.addBreak();
            run1302.setText(" имел(а) _______________________________________________________________________________");
            run1302.addBreak();
            run1302.setText(" ______________________________________________________________________________________");
            run1302.addBreak();

            XWPFRun run1303 = parag13.createRun();
            run1303.setFontFamily("Times New Roman");
            run1303.setFontSize(12);
            run1303.setBold(true);
            run1303.setText("35. Результаты дополнительных исследований:");
            run1303.addBreak();
            XWPFRun run1304 = parag13.createRun();
            run1304.setFontFamily("Times New Roman");
            run1304.setFontSize(12);
            run1304.setText(" ______________________________________________________________________________________");
            run1304.addBreak();
            run1304.setText(" ______________________________________________________________________________________");
            run1304.addBreak();
            run1304.setText(" ______________________________________________________________________________________");
            run1304.addBreak();
            run1304.setText(" ______________________________________________________________________________________");
            run1304.addBreak();
            run1304.setText(" ______________________________________________________________________________________");
            run1304.addBreak();

            XWPFRun run1305 = parag13.createRun();
            run1305.setFontFamily("Times New Roman");
            run1305.setFontSize(12);
            run1305.setBold(true);
            run1305.setText("36. Объективный статус");
            XWPFRun run1306 = parag13.createRun();
            run1306.setFontFamily("Times New Roman");
            run1306.setFontSize(12);
            run1306.setText(" (удовлетворительное, относительно удовл., неудов.): ");
            run1306.addBreak();
            run1306.setText("t" + "°" + "C____________________________________________________________________________________");
            run1306.addBreak();
            run1306.setText("Кожа и видимые слизистые (физиологической окраски, акроцианоз, сыпь)");
            run1306.addBreak();
            run1306.setText(" ______________________________________________________________________________________");
            run1306.addBreak();
            run1306.setText("Состояние зева (не гиперемирован, гиперемирован)");
            run1306.setText(" __________________________________________");
            run1306.addBreak();
            run1306.setText("Тоны сердца (ритмичные, аритмичные, тоны звучные,приглушены,шумы)");
            run1306.setText("  ______________________");
            run1306.addBreak();
            run1306.setText("Акцент ________________________ АД ______________мм рт. ст. ЧСС _________________________");
            run1306.addBreak();
            run1306.setText("В легких дыхание везикулярное, хрипы (не выслушиваются, выслушиваются) ");
            run1306.addBreak();
            run1306.setText(" ______________________________________________________________________________________");
            run1306.addBreak();
            run1306.setText("Мочеиспускание (безболезненное, болезнен.), диурез (адекватный, контролирует, не контр.)");
            run1306.addBreak();
            run1306.setText("Стул оформленный (неоформленный) _______________  Отеки ________________________________");
            run1306.addBreak();
            run1306.setText("Передвигается (самостоятельно,с помощью)");
            run1306.addBreak();
            run1306.setText(" ______________________________________________________________________________________");
            run1306.addBreak();
            run1306.addBreak();
            XWPFRun run1307 = parag13.createRun();
            run1307.setFontFamily("Times New Roman");
            run1307.setFontSize(12);
            run1307.setBold(true);
            run1307.setText("37. Функциональное состояние при поступлении и при выписке");
            XWPFRun run1308 = parag13.createRun();
            run1308.setFontFamily("Times New Roman");
            run1308.setFontSize(12);
            run1308.setText(" (в баллах):");


            XWPFTable table = myNewDoc.createTable();
            table.setWidth(10000);

            //create first row
            XWPFTableRow tableRowOne = table.getRow(0);


            XWPFRun runt1 =  tableRowOne.getCell(0).addParagraph().createRun();
            runt1.setFontFamily("Times New Roman"); runt1.setFontSize(12); runt1.setText(" № п/п ");
            tableRowOne.getCell(0).removeParagraph(0);

            XWPFRun runt2 = tableRowOne.addNewTableCell().addParagraph().createRun();
            runt2.setFontFamily("Times New Roman"); runt2.setFontSize(12); runt2.setText(" Оцениваемая категория ");
            tableRowOne.getCell(1).removeParagraph(0);

            XWPFRun runt3 = tableRowOne.addNewTableCell().addParagraph().createRun();
            runt3.setFontFamily("Times New Roman"); runt3.setFontSize(12); runt3.setText(" При поступлении ");
            tableRowOne.getCell(2).removeParagraph(0);

            XWPFRun runt4 = tableRowOne.addNewTableCell().addParagraph().createRun();
            runt4.setFontFamily("Times New Roman"); runt4.setFontSize(12); runt4.setText(" При выписке ");
            tableRowOne.getCell(3).removeParagraph(0);


            //create second row
            XWPFTableRow tableRowTwo = table.createRow();

            XWPFRun runt01 =  tableRowTwo.getCell(0).addParagraph().createRun();
            runt01.setFontFamily("Times New Roman"); runt01.setFontSize(12); runt01.setText(" 1 ");
            tableRowTwo.getCell(0).removeParagraph(0);

            XWPFRun runt02 = tableRowTwo.getCell(1).addParagraph().createRun();
            runt02.setFontFamily("Times New Roman"); runt02.setFontSize(12); runt02.setText(" Объем активных движений в суставах ");
            tableRowTwo.getCell(1).removeParagraph(0);

            XWPFRun runt03 = tableRowTwo.getCell(2).addParagraph().createRun();
            runt03.setFontFamily("Times New Roman"); runt03.setFontSize(12); runt03.setText("");
            tableRowTwo.getCell(2).removeParagraph(0);

            XWPFRun runt04 = tableRowTwo.getCell(3).addParagraph().createRun();
            runt04.setFontFamily("Times New Roman"); runt04.setFontSize(12); runt04.setText("");
            tableRowTwo.getCell(3).removeParagraph(0);


            //create third row
            XWPFTableRow tableRowThree = table.createRow();

            XWPFRun runt05 =  tableRowThree.getCell(0).addParagraph().createRun();
            runt05.setFontFamily("Times New Roman"); runt05.setFontSize(12); runt05.setText(" 2 ");
            tableRowThree.getCell(0).removeParagraph(0);

            XWPFRun runt06 =  tableRowThree.getCell(1).addParagraph().createRun();
            runt06.setFontFamily("Times New Roman"); runt06.setFontSize(12); runt06.setText(" Мышечная сила ");
            tableRowThree.getCell(1).removeParagraph(0);

            XWPFRun runt07 =  tableRowThree.getCell(2).addParagraph().createRun();
            runt07.setFontFamily("Times New Roman"); runt07.setFontSize(12); runt07.setText("");
            tableRowThree.getCell(2).removeParagraph(0);

            XWPFRun runt08 =  tableRowThree.getCell(3).addParagraph().createRun();
            runt08.setFontFamily("Times New Roman"); runt08.setFontSize(12); runt08.setText("");
            tableRowThree.getCell(3).removeParagraph(0);



            XWPFTableRow tableRowFour = table.createRow();

            XWPFRun runt09 =  tableRowFour.getCell(0).addParagraph().createRun();
            runt09.setFontFamily("Times New Roman"); runt09.setFontSize(12); runt09.setText(" 3 ");
            tableRowFour.getCell(0).removeParagraph(0);

            XWPFRun runt10 =  tableRowFour.getCell(1).addParagraph().createRun();
            runt10.setFontFamily("Times New Roman"); runt10.setFontSize(12); runt10.setText(" Тонус мышц ");
            tableRowFour.getCell(1).removeParagraph(0);

            XWPFRun runt11 =  tableRowFour.getCell(2).addParagraph().createRun();
            runt11.setFontFamily("Times New Roman"); runt11.setFontSize(12); runt11.setText("");
            tableRowFour.getCell(2).removeParagraph(0);

            XWPFRun runt12 =  tableRowFour.getCell(3).addParagraph().createRun();
            runt12.setFontFamily("Times New Roman"); runt12.setFontSize(12); runt12.setText("");
            tableRowFour.getCell(3).removeParagraph(0);


            XWPFTableRow tableRowFive = table.createRow();

            XWPFRun runt13 =  tableRowFive.getCell(0).addParagraph().createRun();
            runt13.setFontFamily("Times New Roman"); runt13.setFontSize(12); runt13.setText(" 4 ");
            tableRowFive.getCell(0).removeParagraph(0);

            XWPFRun runt14 =  tableRowFive.getCell(1).addParagraph().createRun();
            runt14.setFontFamily("Times New Roman"); runt14.setFontSize(12); runt14.setText(" Координация движений ");
            tableRowFive.getCell(1).removeParagraph(0);

            XWPFRun runt15 =  tableRowFive.getCell(2).addParagraph().createRun();
            runt15.setFontFamily("Times New Roman"); runt15.setFontSize(12); runt15.setText("");
            tableRowFive.getCell(2).removeParagraph(0);

            XWPFRun runt16 =  tableRowFive.getCell(3).addParagraph().createRun();
            runt16.setFontFamily("Times New Roman"); runt16.setFontSize(12); runt16.setText("");
            tableRowFive.getCell(3).removeParagraph(0);

            XWPFTableRow tableRowSix = table.createRow();

            XWPFRun runt17 =  tableRowSix.getCell(0).addParagraph().createRun();
            runt17.setFontFamily("Times New Roman"); runt17.setFontSize(12); runt17.setText(" 5 ");
            tableRowSix.getCell(0).removeParagraph(0);

            XWPFRun runt18 =  tableRowSix.getCell(1).addParagraph().createRun();
            runt18.setFontFamily("Times New Roman"); runt18.setFontSize(12); runt18.setText(" Степень самообслуживания ");
            tableRowSix.getCell(1).removeParagraph(0);

            XWPFRun runt19 =  tableRowSix.getCell(2).addParagraph().createRun();
            runt19.setFontFamily("Times New Roman"); runt19.setFontSize(12); runt19.setText("");
            tableRowSix.getCell(2).removeParagraph(0);

            XWPFRun runt20 =  tableRowSix.getCell(3).addParagraph().createRun();
            runt20.setFontFamily("Times New Roman"); runt20.setFontSize(12); runt20.setText("");
            tableRowSix.getCell(3).removeParagraph(0);

            // Конец таблицы

            XWPFParagraph parag014 = myNewDoc.createParagraph();

            parag014.setAlignment(ParagraphAlignment.LEFT);
            parag014.setPageBreak(true);
            parag014.setSpacingBetween(1.25);
            parag014.setSpacingAfterLines(0);
            XWPFRun run1401 = parag014.createRun();
            run1401.setFontFamily("Times New Roman");
            run1401.setFontSize(12);
            run1401.setBold(true);
            run1401.setText("38. Заключение врача ");


            XWPFParagraph parag14 = myNewDoc.createParagraph();

            parag14.setAlignment(ParagraphAlignment.BOTH);
            parag14.setSpacingBetween(1.25);
            parag14.setSpacingAfterLines(0);

            XWPFRun run1402 = parag14.createRun();
            run1402.setFontFamily("Times New Roman");
            run1402.setFontSize(14);
            run1402.setUnderline(UnderlinePatterns.SINGLE);
            run1402.setText("На основании данных объективного  осмотра и представленной медицинской документации, рекомендовано проведение восстановительной терапии и социально-реабилитационные мероприятия.");

            XWPFRun run1403 = parag14.createRun();
            run1403.setFontFamily("Times New Roman");
            run1403.setFontSize(14);
            run1403.setText("______________________________________________");
            run1403.addBreak();
            run1403.setText("__________________________________________________________________________");
            run1403.addBreak();
            run1403.setText("__________________________________________________________________________");
            run1403.addBreak();
            run1403.setText("__________________________________________________________________________");
            run1403.addBreak();
            run1403.setText("__________________________________________________________________________");

            XWPFParagraph parag15 = myNewDoc.createParagraph();

            parag15.setAlignment(ParagraphAlignment.LEFT);
            parag15.setSpacingBetween(1.25);
            parag15.setSpacingAfterLines(0);

            XWPFRun run1501 = parag15.createRun();
            run1501.setFontFamily("Times New Roman");
            run1501.setFontSize(14);
            run1501.setBold(true);
            run1501.setText("Врач-терапевт   ________________   _____________________________ ");

            XWPFParagraph parag16 = myNewDoc.createParagraph();

            parag16.setAlignment(ParagraphAlignment.CENTER);
            parag16.setSpacingBetween(1.25);
            parag16.setSpacingAfterLines(0);

            XWPFRun run1601 = parag16.createRun();
            run1601.setFontFamily("Times New Roman");
            run1601.setFontSize(12);
            run1601.setBold(true);
            run1601.setText("(подпись)        (расшифровка подписи, Ф.И.О) ");
            run1601.addBreak();

            XWPFParagraph parag17 = myNewDoc.createParagraph();

            parag17.setAlignment(ParagraphAlignment.CENTER);
            parag17.setSpacingBetween(1.25);
            parag17.setSpacingAfterLines(0);

            XWPFRun run1701 = parag17.createRun();
            run1701.setFontFamily("Times New Roman");
            run1701.setFontSize(12);
            run1701.setBold(true);
            run1701.setText("ЛИСТ НАЗНАЧЕНИЙ");
// Таблица на стр №5
            XWPFTable table1 = myNewDoc.createTable();

            XWPFTableRow table1RowOne = table1.getRow(0);



            XWPFRun run1t1 =  table1RowOne.getCell(0).addParagraph().createRun();
            run1t1.setFontFamily("Times New Roman"); run1t1.setFontSize(12); run1t1.setText(" Дата ");
            table1RowOne.getCell(0).removeParagraph(0);

            XWPFRun run1t2 = table1RowOne.addNewTableCell().addParagraph().createRun();
            run1t2.setFontFamily("Times New Roman"); run1t2.setFontSize(12); run1t2.setText(" Назначение ");
            table1RowOne.getCell(1).removeParagraph(0);

            XWPFRun run1t3 = table1RowOne.addNewTableCell().addParagraph().createRun();
            run1t3.setFontFamily("Times New Roman"); run1t3.setFontSize(12); run1t3.setText(" Особые отметки ");
            table1RowOne.getCell(2).removeParagraph(0);

            table1RowOne.getCell(0).setWidth("1000");
            table1RowOne.getCell(1).setWidth("8000");
            table1RowOne.getCell(2).setWidth("4000");
            table1RowOne.setHeight(300);
            table1RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);


            for (int i = 0; i < 28; i++) {
                String s1 = "";
                if (i == 0) {
                    s1 = "1. Диета";
                } else
                if (i == 1) {
                    s1 = "2. ЛФК";
                } else
                if (i == 2) {
                    s1 = "3. Консультация физиотерапевта";
                } else {
                    s1 = " ";
                }
                XWPFTableRow table1RowTwo = table1.createRow();
                table1RowTwo.setHeight(300);
                table1RowTwo.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

                XWPFRun run1t01 = table1RowTwo.getCell(0).addParagraph().createRun();
                run1t01.setFontFamily("Times New Roman");
                run1t01.setFontSize(12);
                run1t01.setText(" ");
                table1RowTwo.getCell(0).removeParagraph(0);

                XWPFRun run1t02 = table1RowTwo.getCell(1).addParagraph().createRun();
                run1t02.setFontFamily("Times New Roman");
                run1t02.setFontSize(12);
                run1t02.setText(s1);
                table1RowTwo.getCell(1).removeParagraph(0);

                XWPFRun run1t03 = table1RowTwo.getCell(2).addParagraph().createRun();
                run1t03.setFontFamily("Times New Roman");
                run1t03.setFontSize(12);
                run1t03.setText(" ");
                table1RowTwo.getCell(2).removeParagraph(0);

            }
            // СТРАНИЦА 6

            XWPFParagraph parag18 = myNewDoc.createParagraph();

            parag18.setAlignment(ParagraphAlignment.CENTER);
            parag18.setPageBreak(true);
            parag18.setSpacingBetween(1.25);
            parag18.setSpacingAfterLines(0);

            XWPFRun run1801 = parag18.createRun();
            run1801.setFontFamily("Times New Roman");
            run1801.setFontSize(12);
            run1801.setBold(true);
            run1801.setText("1.1 РЕАБИЛИТАЦИОННЫЕ МЕРОПРИЯТИЯ МЕДИЦИНСКОГО ХАРАКТЕРА (ВОССТАНОВИТЕЛЬНАЯ ТЕРАПИЯ)");
            run1801.addBreak();
            run1801.addBreak();
            run1801.setText("ФИЗИЧЕСКАЯ РЕАБИЛИТАЦИЯ");

            XWPFParagraph parag19 = myNewDoc.createParagraph();

            parag19.setAlignment(ParagraphAlignment.LEFT);
            parag19.setSpacingBetween(1.25);
            parag19.setSpacingAfterLines(0);

            XWPFRun run1901 = parag19.createRun();
            run1901.setFontFamily("Times New Roman");
            run1901.setFontSize(12);
            run1901.setBold(true);
            run1901.setText("1.1.1 Дата ____________________");
            run1901.addBreak();
            run1901.addBreak();
            run1901.setText("1.1.2 Заключение о возможности проведения реабилитации");
            XWPFParagraph parag20 = myNewDoc.createParagraph();
            parag20.setAlignment(ParagraphAlignment.BOTH);
            parag20.setSpacingBetween(1.25);
            parag20.setSpacingAfterLines(0);
            parag20.setIndentationFirstLine(700);
            XWPFRun run2002 = parag20.createRun();
            run2002.setFontFamily("Times New Roman");
            run2002.setFontSize(12);
            run2002.setUnderline(UnderlinePatterns.SINGLE);
            run2002.setText("На основании медицинского заключения и объективного осмотра показано проведение реабилитационных мероприятий медицинского характера-восстановительной терапии ");
            XWPFRun run2003 = parag20.createRun();
            run2003.setFontFamily("Times New Roman");
            run2003.setFontSize(12);
            run2003.setText("______");
            run2003.addBreak();
            run2003.setText("_______________________________________________________________________________________");
            run2003.addBreak();


            XWPFParagraph parag21 = myNewDoc.createParagraph();

            parag21.setAlignment(ParagraphAlignment.LEFT);
            parag21.setSpacingBetween(1.25);
            parag21.setSpacingAfterLines(0);

            XWPFRun run2101 = parag21.createRun();
            run2101.setFontFamily("Times New Roman");
            run2101.setFontSize(12);
            run2101.setBold(true);
            run2101.setText("1.1.3 Цель физической реабилитации");
            XWPFParagraph parag22 = myNewDoc.createParagraph();
            parag22.setAlignment(ParagraphAlignment.BOTH);
            parag22.setSpacingBetween(1.25);
            parag22.setSpacingAfterLines(0);
            parag22.setIndentationFirstLine(700);
            XWPFRun run2201 = parag22.createRun();
            run2201.setFontFamily("Times New Roman");
            run2201.setFontSize(12);
            run2201.setUnderline(UnderlinePatterns.SINGLE);
            run2201.setText("Снижение болевого синдрома, увеличение объема движений в суставах конечностей, улучшение качества жизни, улучшение статодинамической функции ");
            XWPFRun run2203 = parag22.createRun();
            run2203.setFontFamily("Times New Roman");
            run2203.setFontSize(12);
            run2203.setText("______________________");
            run2203.addBreak();
            run2203.setText("_______________________________________________________________________________________");
            run2203.addBreak();

            XWPFParagraph parag23 = myNewDoc.createParagraph();

            parag23.setAlignment(ParagraphAlignment.LEFT);
            parag23.setSpacingBetween(1.25);
            parag23.setSpacingAfterLines(0);

            XWPFRun run2301 = parag23.createRun();
            run2301.setFontFamily("Times New Roman");
            run2301.setFontSize(12);
            run2301.setBold(true);
            run2301.setText("1.1.4 Программа физической реабилитации");
            run2301.addBreak();
            run2301.addBreak();
            run2301.setText("1.1.5 Кинезотерапия ");
            XWPFRun run2302 = parag23.createRun();
            run2302.setFontFamily("Times New Roman");
            run2302.setFontSize(12);
            run2302.setText("(далее КТ)");


            XWPFParagraph parag24 = myNewDoc.createParagraph();
            parag24.setAlignment(ParagraphAlignment.BOTH);
            parag24.setSpacingBetween(1.25);
            parag24.setSpacingAfterLines(0);
            parag24.setIndentationFirstLine(700);

            XWPFRun run2401 = parag24.createRun();
            run2401.setFontFamily("Times New Roman");
            run2401.setFontSize(12);
            run2401.setText("- по анатомическому типу: ");
            XWPFRun run2402 = parag24.createRun();
            run2402.setFontFamily("Times New Roman");
            run2402.setFontSize(12);
            run2402.setUnderline(UnderlinePatterns.SINGLE);
            run2402.setText("упражнения для мышц позвоночника, мышц брюшного пресса, мышц конечностей ");
            XWPFRun run2403 = parag24.createRun();
            run2403.setFontFamily("Times New Roman");
            run2403.setFontSize(12);
            run2403.setText("______________________________________________________________________");
            XWPFParagraph parag25 = myNewDoc.createParagraph();
            parag25.setAlignment(ParagraphAlignment.BOTH);
            parag25.setSpacingBetween(1.25);
            parag25.setSpacingAfterLines(0);
            parag25.setIndentationFirstLine(700);
            XWPFRun run2501 = parag25.createRun();
            run2501.setFontFamily("Times New Roman");
            run2501.setFontSize(12);
            run2501.setText("- по активности: ");
            XWPFRun run2502 = parag25.createRun();
            run2502.setFontFamily("Times New Roman");
            run2502.setFontSize(12);
            run2502.setUnderline(UnderlinePatterns.SINGLE);
            run2502.setText("активные ");
            XWPFRun run2503 = parag25.createRun();
            run2503.setFontFamily("Times New Roman");
            run2503.setFontSize(12);
            run2503.setText("__________________________________________________________");
            run2503.addBreak();
            run2503.setText("_______________________________________________________________________________________");

            XWPFParagraph parag26 = myNewDoc.createParagraph();
            parag26.setAlignment(ParagraphAlignment.BOTH);
            parag26.setSpacingBetween(1.25);
            parag26.setSpacingAfterLines(0);
            parag26.setIndentationFirstLine(700);
            XWPFRun run2601 = parag26.createRun();
            run2601.setFontFamily("Times New Roman");
            run2601.setFontSize(12);
            run2601.setText("- по использованию снарядов: ");
            XWPFRun run2602 = parag26.createRun();
            run2602.setFontFamily("Times New Roman");
            run2602.setFontSize(12);
            run2602.setUnderline(UnderlinePatterns.SINGLE);
            run2602.setText("с использованием беговой дорожки, велотренажера, гребного тренажера ");
            XWPFRun run2603 = parag26.createRun();
            run2603.setFontFamily("Times New Roman");
            run2603.setFontSize(12);
            run2603.setText(" _____________________________________________________________________________");

            XWPFParagraph parag27 = myNewDoc.createParagraph();
            parag27.setAlignment(ParagraphAlignment.BOTH);
            parag27.setSpacingBetween(1.25);
            parag27.setSpacingAfterLines(0);
            parag27.setIndentationFirstLine(700);
            XWPFRun run2701 = parag27.createRun();
            run2701.setFontFamily("Times New Roman");
            run2701.setFontSize(12);
            run2701.setText("- по видовому признаку: ");
            XWPFRun run2702 = parag27.createRun();
            run2702.setFontFamily("Times New Roman");
            run2702.setFontSize(12);
            run2702.setUnderline(UnderlinePatterns.SINGLE);
            run2702.setText("статические, динамические, на улучшение подвижности суставов ");
            run2702.addBreak();
            XWPFRun run2703 = parag27.createRun();
            run2703.setFontFamily("Times New Roman");
            run2703.setFontSize(12);
            run2703.setText("_______________________________________________________________________________________");
            run2703.addBreak();
            run2703.setText("_______________________________________________________________________________________");
            run2703.addBreak();

            XWPFParagraph parag28 = myNewDoc.createParagraph();

            parag28.setAlignment(ParagraphAlignment.LEFT);
            parag28.setSpacingBetween(1.25);
            parag28.setSpacingAfterLines(0);

            XWPFRun run2801 = parag28.createRun();
            run2801.setFontFamily("Times New Roman");
            run2801.setFontSize(12);
            run2801.setBold(true);
            run2801.setText("1.1.6 Лечебная физическая культура");
            XWPFRun run2802 = parag28.createRun();
            run2802.setFontFamily("Times New Roman");
            run2802.setFontSize(12);
            run2802.setText(" (далее ЛФК):");

            XWPFParagraph parag29 = myNewDoc.createParagraph();
            parag29.setAlignment(ParagraphAlignment.BOTH);
            parag29.setSpacingBetween(1.25);
            parag29.setSpacingAfterLines(0);
            parag29.setIndentationFirstLine(700);
            XWPFRun run2901 = parag29.createRun();
            run2901.setFontFamily("Times New Roman");
            run2901.setFontSize(12);
            run2901.setText("- физические (гимнастические) упражнения (общеразвивающие, специальные): ");
            XWPFRun run2902 = parag29.createRun();
            run2902.setFontFamily("Times New Roman");
            run2902.setFontSize(12);
            run2902.setUnderline(UnderlinePatterns.SINGLE);
            run2902.setText("на увеличение объема движений в суставах конечностей, укрепление мышц конечностей, коррекцию осанки ");
            run2902.addBreak();
            XWPFRun run2903 = parag29.createRun();
            run2903.setFontFamily("Times New Roman");
            run2903.setFontSize(12);
            run2903.setText("_______________________________________________________________________________________");

            XWPFParagraph parag30 = myNewDoc.createParagraph();
            parag30.setAlignment(ParagraphAlignment.BOTH);
            parag30.setSpacingBetween(1.25);
            parag30.setSpacingAfterLines(0);
            parag30.setIndentationFirstLine(700);
            XWPFRun run3001 = parag30.createRun();
            run3001.setFontFamily("Times New Roman");
            run3001.setFontSize(12);
            run3001.setText("- спортивно-прикладные упражнения: ");
            XWPFRun run3002 = parag30.createRun();
            run3002.setFontFamily("Times New Roman");
            run3002.setFontSize(12);
            run3002.setUnderline(UnderlinePatterns.SINGLE);
            run3002.setText("с использованием гимнастической стенки, скамейки; скандинавская ходьба ");
            XWPFRun run3003 = parag30.createRun();
            run3003.setFontFamily("Times New Roman");
            run3003.setFontSize(12);
            run3003.setText(" ___________________________________________________________________");


            // СТРАНИЦА 7

            XWPFParagraph parag31 = myNewDoc.createParagraph();

            parag31.setAlignment(ParagraphAlignment.LEFT);
            parag31.setSpacingBetween(1.25);
            parag31.setSpacingAfterLines(0);
            parag31.setPageBreak(true);

            XWPFRun run3101 = parag31.createRun();
            run3101.setFontFamily("Times New Roman");
            run3101.setFontSize(12);
            run3101.setBold(true);
            run3101.setText("1.1.7 Механотерапия ");
            XWPFRun run3102 = parag31.createRun();
            run3102.setFontFamily("Times New Roman");
            run3102.setFontSize(12);
            run3102.setText(" (далее МТ):");

            XWPFParagraph parag32 = myNewDoc.createParagraph();
            parag32.setAlignment(ParagraphAlignment.BOTH);
            parag32.setSpacingBetween(1.25);
            parag32.setSpacingAfterLines(0);
            parag32.setIndentationFirstLine(700);
            XWPFRun run3202 = parag32.createRun();
            run3202.setFontFamily("Times New Roman");
            run3202.setFontSize(12);
            run3202.setUnderline(UnderlinePatterns.SINGLE);
            run3202.setText("Занятия на блоковых тренажерах, беговой дорожке, велотренажере, гребном тренажере, гидротерапия ");
            XWPFRun run3203 = parag32.createRun();
            run3203.setFontFamily("Times New Roman");
            run3203.setFontSize(12);
            run3203.setText(" __________________________________________________________________________");
            run3203.addBreak();
            run3203.setText("_______________________________________________________________________________________");
            run3203.addBreak();
            run3203.setText("_______________________________________________________________________________________");
            run3203.addBreak();

            XWPFParagraph parag33 = myNewDoc.createParagraph();

            parag33.setAlignment(ParagraphAlignment.LEFT);
            parag33.setSpacingBetween(1.25);
            parag33.setSpacingAfterLines(0);

            XWPFRun run3301 = parag33.createRun();
            run3301.setFontFamily("Times New Roman");
            run3301.setFontSize(12);
            run3301.setBold(true);
            run3301.setText("1.1.8 Лечебный массаж ");

            XWPFParagraph parag34 = myNewDoc.createParagraph();
            parag34.setAlignment(ParagraphAlignment.BOTH);
            parag34.setSpacingBetween(1.25);
            parag34.setSpacingAfterLines(0);
            parag34.setIndentationFirstLine(700);
            XWPFRun run3402 = parag34.createRun();
            run3402.setFontFamily("Times New Roman");
            run3402.setFontSize(12);
            run3402.setText("Вид массажа, область воздействия: ");
            XWPFRun run3403 = parag34.createRun();
            run3403.setFontFamily("Times New Roman");
            run3403.setFontSize(12);
            run3403.setUnderline(UnderlinePatterns.SINGLE);
            run3403.setText(" общеукрепляющий, оздоровительный массаж");
            XWPFRun run3404 = parag34.createRun();
            run3404.setFontFamily("Times New Roman");
            run3404.setFontSize(12);
            run3404.setText(" шейно-воротниковой зоны грудного отдела позвоночника, пояснично-крестцового отдела позвоночника, верхних конечностей, нижних конечностей\t");
            run3404.addBreak();

            XWPFParagraph parag35 = myNewDoc.createParagraph();

            parag35.setAlignment(ParagraphAlignment.LEFT);
            parag35.setSpacingBetween(1.25);
            parag35.setSpacingAfterLines(0);

            XWPFRun run3501 = parag35.createRun();
            run3501.setFontFamily("Times New Roman");
            run3501.setFontSize(12);
            run3501.setBold(true);
            run3501.setText("1.1.9 Физиотерапия ");

            XWPFParagraph parag36 = myNewDoc.createParagraph();
            parag36.setAlignment(ParagraphAlignment.BOTH);
            parag36.setSpacingBetween(1.25);
            parag36.setSpacingAfterLines(0);
            parag36.setIndentationFirstLine(700);
            XWPFRun run3601 = parag36.createRun();
            run3601.setFontFamily("Times New Roman");
            run3601.setFontSize(12);
            run3601.setText("Вид процедуры, область воздействия: \t");
            run3601.addBreak();
            run3601.setText("_______________________________________________________________________________________");
            run3601.addBreak();
            run3601.setText("_______________________________________________________________________________________");
            run3601.addBreak();
            run3601.setText("_______________________________________________________________________________________");
            run3601.addBreak();
            run3601.setText("_______________________________________________________________________________________");
            run3601.addBreak();

            XWPFParagraph parag37 = myNewDoc.createParagraph();
            parag37.setAlignment(ParagraphAlignment.LEFT);
            parag37.setSpacingBetween(1.25);
            parag37.setSpacingAfterLines(0);
            XWPFRun run3701 = parag37.createRun();
            run3701.setFontFamily("Times New Roman");
            run3701.setFontSize(12);
            run3701.setBold(true);
            run3701.setText("Итоги проведенных мероприятий по восстановительной терапии:");

            XWPFTable table2 = myNewDoc.createTable();
            table2.setWidth(10000);

            //create first row
            XWPFTableRow table2RowOne = table2.getRow(0);

            XWPFRun run2t1 =  table2RowOne.getCell(0).addParagraph().createRun();
            run2t1.setFontFamily("Times New Roman"); run2t1.setFontSize(12); run2t1.setText(" Наименование услуги ");
            table2RowOne.getCell(0).removeParagraph(0);

            XWPFRun run2t2 = table2RowOne.addNewTableCell().addParagraph().createRun();
            run2t2.setFontFamily("Times New Roman"); run2t2.setFontSize(12); run2t2.setText(" ЛФК ");
            table2RowOne.getCell(1).removeParagraph(0);

            XWPFRun run2t3 = table2RowOne.addNewTableCell().addParagraph().createRun();
            run2t3.setFontFamily("Times New Roman"); run2t3.setFontSize(12); run2t3.setText(" МТ ");
            table2RowOne.getCell(2).removeParagraph(0);

            XWPFRun run2t4 = table2RowOne.addNewTableCell().addParagraph().createRun();
            run2t4.setFontFamily("Times New Roman"); run2t4.setFontSize(12); run2t4.setText(" КТ ");
            table2RowOne.getCell(3).removeParagraph(0);

            XWPFRun run2t5 = table2RowOne.addNewTableCell().addParagraph().createRun();
            run2t5.setFontFamily("Times New Roman"); run2t5.setFontSize(12); run2t5.setText(" Массаж ");
            table2RowOne.getCell(4).removeParagraph(0);

            XWPFRun run2t6 = table2RowOne.addNewTableCell().addParagraph().createRun();
            run2t6.setFontFamily("Times New Roman"); run2t6.setFontSize(12); run2t6.setText(" Физиотерапия ");
            table2RowOne.getCell(5).removeParagraph(0);


            //create second row
            XWPFTableRow table2RowTwo = table2.createRow();

            XWPFRun run2t01 =  table2RowTwo.getCell(0).addParagraph().createRun();
            run2t01.setFontFamily("Times New Roman"); run2t01.setFontSize(12); run2t01.setText(" Количество услуг ");
            table2RowTwo.getCell(0).removeParagraph(0);


            XWPFParagraph parag38 = myNewDoc.createParagraph();
            parag38.setAlignment(ParagraphAlignment.LEFT);
            parag38.setSpacingBetween(1.25);
            parag38.setSpacingAfterLines(0);

            XWPFRun run3801 = parag38.createRun();
            run3801.setFontFamily("Times New Roman");
            run3801.setFontSize(12);
            run3801.setBold(true);
            run3801.addBreak();
            run3801.setText("1.2 Заключение по итогам реабилитационных мероприятий по восстановительной терапии: ");

            XWPFParagraph parag39 = myNewDoc.createParagraph();
            parag39.setAlignment(ParagraphAlignment.BOTH);
            parag39.setSpacingBetween(1.25);
            parag39.setSpacingAfterLines(0);
            parag39.setIndentationFirstLine(700);
            XWPFRun run3902 = parag39.createRun();
            run3902.setFontFamily("Times New Roman");
            run3902.setFontSize(12);
            run3902.setUnderline(UnderlinePatterns.SINGLE);
            run3902.setText("На фоне проведения восстановительной терапии отмечается снижение болевого синдрома, увеличение объема движений в суставах конечностей, улучшение статодинамической функции. Качество жизни значительно улучшилось\t");
            XWPFRun run3903 = parag39.createRun();
            run3903.setFontFamily("Times New Roman");
            run3903.setFontSize(12);
            run3903.setText(" _____________________________________________");
            run3903.addBreak();
            run3903.setText("_______________________________________________________________________________________");
            run3903.addBreak();
            run3903.setText("_______________________________________________________________________________________");
            run3903.addBreak();
            run3903.addBreak();
            run3903.setText("Врач-реабилитолог________________________________________________    Ю.О. Щербаков\t");

            // СТРАНИЦА 8
            XWPFParagraph parag40 = myNewDoc.createParagraph();

            parag40.setAlignment(ParagraphAlignment.CENTER);
            parag40.setSpacingBetween(1.05);
            parag40.setSpacingAfterLines(0);
            parag40.setPageBreak(true);

            XWPFRun run4001 = parag40.createRun();
            run4001.setFontFamily("Times New Roman");
            run4001.setFontSize(13);
            run4001.setBold(true);
            run4001.setText("IV. СОЦИАЛЬНАЯ РЕАБИЛИТАЦИЯ");
            run4001.addBreak();

            XWPFParagraph parag41 = myNewDoc.createParagraph();

            parag41.setAlignment(ParagraphAlignment.LEFT);
            parag41.setSpacingBetween(1.05);
            parag41.setSpacingAfterLines(0);

            XWPFRun run4101 = parag41.createRun();
            run4101.setFontFamily("Times New Roman");
            run4101.setFontSize(13);
            run4101.setBold(true);
            run4101.setText("Дата ____________________");
            run4101.addBreak();
            run4101.addBreak();
            run4101.setText("4.1. СОЦИАЛЬНО-СРЕДОВАЯ РЕАБИЛИТАЦИЯ");
            run4101.addBreak();
            run4101.setText("4.1.1. Оценка нарушений социально-средового статуса: ");
            XWPFRun run4102 = parag41.createRun();
            run4102.setFontFamily("Times New Roman");
            run4102.setFontSize(12);
            run4102.setText("без нарушений, легкие нарушения, средние нарушения, выраженные нарушения");
            run4102.addBreak();
            XWPFRun run4103 = parag41.createRun();
            run4103.setFontFamily("Times New Roman");
            run4103.setFontSize(13);
            run4103.setBold(true);
            run4103.setText("4.1.2.  Программа социально-средовой реабилитации:");
            run4103.addBreak();
            run4103.setText("4.1.2.1. Обучение инвалида и членов его семьи пользованию техническими средствами реабилитации ");
            XWPFRun run4104 = parag41.createRun();
            run4104.setFontFamily("Times New Roman");
            run4104.setFontSize(13);
            run4104.setItalic(true);
            run4104.setBold(true);
            run4104.setUnderline(UnderlinePatterns.SINGLE);
            run4104.setText("(рекомендациии)");
            XWPFRun run4105 = parag41.createRun();
            run4105.setFontFamily("Times New Roman");
            run4105.setFontSize(12);
            run4105.setText(": _______________________________________________________");
            run4105.addBreak();
            run4105.setText("_______________________________________________________________________________________");
            run4105.addBreak();
            run4105.setText("_______________________________________________________________________________________");
            run4105.addBreak();
            run4105.setText("_______________________________________________________________________________________");
            run4105.addBreak();
            XWPFRun run4106 = parag41.createRun();
            run4106.setFontFamily("Times New Roman");
            run4106.setFontSize(13);
            run4106.setBold(true);
            run4106.setText("4.1.2.2. Рекомендации по адаптации жилья к потребностям инвалида: ");
            run4106.addBreak();
            XWPFRun run4107 = parag41.createRun();
            run4107.setFontFamily("Times New Roman");
            run4107.setFontSize(12);
            run4107.setText("_______________________________________________________________________________________");
            run4107.addBreak();
            run4107.setText("_______________________________________________________________________________________");
            run4107.addBreak();
            run4107.setText("_______________________________________________________________________________________");
            run4107.addBreak();
            XWPFRun run4108 = parag41.createRun();
            run4108.setFontFamily("Times New Roman");
            run4108.setFontSize(13);
            run4108.setBold(true);
            run4108.setText("4.1.3. Итог проведенных мероприятий по социально-средовой реабилитации: ");

   //Таблица
            XWPFTable table3 = myNewDoc.createTable();
            table3.setWidth(10700);

            //create first row
            XWPFTableRow table3RowOne = table3.getRow(0);


            XWPFRun run3t1 =  table3RowOne.getCell(0).addParagraph().createRun();
            run3t1.setFontFamily("Times New Roman"); run3t1.setFontSize(12); run3t1.setText(" Наименование ");
            run3t1.setBold(true); run3t1.addBreak(); run3t1.setText(" услуги");
            table3RowOne.getCell(0).removeParagraph(0);

            XWPFRun run3t2 = table3RowOne.addNewTableCell().addParagraph().createRun();
            run3t2.setFontFamily("Times New Roman"); run3t2.setFontSize(12); run3t2.setText(" Обучение инвалида и членов его семьи ");
            run3t2.setBold(true); run3t2.addBreak(); run3t2.setText(" пользованию  техническими средствами ");
            run3t2.addBreak(); run3t2.setText(" реабилитации (ТСР) ");
            table3RowOne.getCell(1).removeParagraph(0);

            XWPFRun run3t3 = table3RowOne.addNewTableCell().addParagraph().createRun();
            run3t3.setFontFamily("Times New Roman"); run3t3.setFontSize(12); run3t3.setText(" Рекомендации по адаптации жилья ");
            run3t3.setBold(true); run3t3.addBreak(); run3t3.setText(" к потребностям инвалида ");
            table3RowOne.getCell(2).removeParagraph(0);

            table3RowOne.getCell(0).setWidth("1000");
            table3RowOne.getCell(1).setWidth("6600");
            table3RowOne.getCell(2).setWidth("5400");

            //create second row
            XWPFTableRow table3RowTwo = table3.createRow();

            XWPFRun run3t01 = table3RowTwo.getCell(0).addParagraph().createRun();
            run3t01.setFontFamily("Times New Roman"); run3t01.setFontSize(12); run3t01.setText(" Количество услуг");
            run3t01.setBold(true);
            table3RowTwo.getCell(0).removeParagraph(0);



            XWPFParagraph parag42 = myNewDoc.createParagraph();

            parag42.setAlignment(ParagraphAlignment.LEFT);
            parag42.setSpacingBetween(1.05);
            parag42.setSpacingAfterLines(0);

            XWPFRun run4201 = parag42.createRun();
            run4201.setFontFamily("Times New Roman");
            run4201.setFontSize(13);
            run4201.setBold(true);
            run4201.addBreak();
            run4201.setText("4.2. СОЦИАЛЬНО-БЫТОВАЯ АДАПТАЦИЯ");
            run4201.addBreak();
            run4201.setText("4.2.1. Оценка нарушений социально-бытового статуса: ");
            XWPFRun run4202 = parag42.createRun();
            run4202.setFontFamily("Times New Roman");
            run4202.setFontSize(13);
            run4202.setText("без нарушений, легкие нарушения, средние нарушения, выраженные нарушения");
            run4202.addBreak();
            XWPFRun run4203 = parag42.createRun();
            run4203.setFontFamily("Times New Roman");
            run4203.setFontSize(13);
            run4203.setBold(true);
            run4203.setText("4.2.2. Программа социально-бытовой адаптации:");
            run4203.addBreak();
            run4203.setText("4.2.2.1. Обучение инвалида навыкам личной гигиены и самообслуживания, в том числе с помощью ТСР:");
            run4203.addBreak();
            XWPFRun run4204 = parag42.createRun();
            run4204.setFontFamily("Times New Roman");
            run4204.setFontSize(12);
            run4204.setBold(true);
            run4204.setText("персональный уход : возможность соблюдения личной гигиены ");
            run4204.setText(" ____________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("возможность пользоваться одеждой ");
            run4204.setText("______________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("возможность приема пищи ");
            run4204.setText("____________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            run4204.addBreak();
            run4204.setText("_______________________________________________________________________________________");
            XWPFParagraph parag43 = myNewDoc.createParagraph();

            parag43.setAlignment(ParagraphAlignment.BOTH);
            parag43.setSpacingBetween(1.05);
            parag43.setSpacingAfterLines(0);
            XWPFRun run4301 = parag43.createRun();
            run4301.setFontFamily("Times New Roman");
            run4301.setFontSize(12);
            run4301.setBold(true);
            run4301.setText("4.2.2.2 обучение пользованию ТСР для обслуживания (далее ТСР) ");
            XWPFRun run4302 = parag43.createRun();
            run4302.setFontFamily("Times New Roman");
            run4302.setFontSize(12);
            run4302.setText("(противоскользящий коврик в душе, мыльница-дозатор, комплект для уборки помещения, приспособление для одевания носков, захват для надевания чулок, приспособление для снятия обуви, активный захват для открывания дверей и окон, активный захват для ключей, активный захват с щипцами на конце, крючок для застегивания пуговиц, приспособление для самостоятельного одевания, устройство для застегивания, захват для надевания рубашек, ограничитель на тарелку, зажим для открывания винтовых крышек, многофункциональная доска, сахарница с дозирующим устройством, перечница с дозирующим устройством, чашка-поильник, стенд для отработки бытовых навыков, куб-застежка, прибор для удержания (ручек, карандашей), приспособление для письма ручкой, прикроватный столик, костыли локтевые, кресло-коляска (инвалидное) прогулочная, кресло-коляска (инвалидное)легкая, обучение письму, обучение надеванию протеза(с помощью протяжки, вакуумное крепление), обучение уходу за культеприемной гильзой) \t");
            run4302.addBreak();
            run4302.setText("_______________________________________________________________________________________");
            run4302.addBreak();
            run4302.setText("_______________________________________________________________________________________");
            run4302.addBreak();
            run4302.setText("_______________________________________________________________________________________");
            XWPFRun run4303 = parag43.createRun();
            run4303.setFontFamily("Times New Roman");
            run4303.setFontSize(12);
            run4303.setBold(true);
            run4303.setText("4.2.2.3. обучение передвижению (далее ОП) ");
            XWPFRun run4304 = parag43.createRun();
            run4304.setFontFamily("Times New Roman");
            run4304.setFontSize(12);
            run4304.setText("(выработка грамотного стереотипа ходьбы с помощью следовой дорожки, подниматься и спускаться по лестнице, передвижение с ходунками, тростью, костылями, ходунки на колесиках, подъем и спуск по пандусу, платформа подъемная вертикальная для инвалидов, передвижение с помощью ходунки-опора, обучение ходьбе на протезе, формирование элементов шага, освоение самостоятельного передвижения и выработки нового стереотипа ходьбы при пользовании протезом) \t");
            run4304.addBreak();
            run4304.setText("_______________________________________________________________________________________");
            run4304.addBreak();
            run4304.setText("_______________________________________________________________________________________");
            run4304.addBreak();
            run4304.setText("_______________________________________________________________________________________");
            run4304.addBreak();
            run4304.addBreak();
            XWPFRun run4305 = parag43.createRun();
            run4305.setFontFamily("Times New Roman");
            run4305.setFontSize(13);
            run4305.setBold(true);
            run4305.setText("4.2.3. Итог проведенных мероприятий по социальной адаптации: \t");
            run4305.addBreak();


            // Таблица
            XWPFTable table4 = myNewDoc.createTable();
            table4.setWidth(10700);

            //create first row
            XWPFTableRow table4RowOne = table4.getRow(0);


            XWPFRun run4t1 =  table4RowOne.getCell(0).addParagraph().createRun();
            run4t1.setFontFamily("Times New Roman"); run4t1.setFontSize(12); run4t1.setText(" Наименование ");
            run4t1.setBold(true); run4t1.addBreak(); run4t1.setText(" услуги");
            table4RowOne.getCell(0).removeParagraph(0);

            XWPFRun run4t2 = table4RowOne.addNewTableCell().addParagraph().createRun();
            run4t2.setFontFamily("Times New Roman"); run4t2.setFontSize(12); run4t2.setText(" Персональный уход ");
            run4t2.setBold(true);
            table4RowOne.getCell(1).removeParagraph(0);

            XWPFRun run4t3 = table4RowOne.addNewTableCell().addParagraph().createRun();
            run4t3.setFontFamily("Times New Roman"); run4t3.setFontSize(12); run4t3.setText(" Обучение ");
            run4t3.setBold(true); run4t3.addBreak(); run4t3.setText(" передвижению ");
            table4RowOne.getCell(2).removeParagraph(0);

            XWPFRun run4t4 = table4RowOne.addNewTableCell().addParagraph().createRun();
            run4t4.setFontFamily("Times New Roman"); run4t4.setFontSize(12); run4t4.setText(" Обучение пользованию ТСР ");
            run4t4.setBold(true); run4t4.addBreak(); run4t4.setText(" для самообслуживания ");
            table4RowOne.getCell(3).removeParagraph(0);

            table4RowOne.getCell(0).setWidth("1000");
            table4RowOne.getCell(1).setWidth("3000");
            table4RowOne.getCell(2).setWidth("3000");
            table4RowOne.getCell(3).setWidth("5000");

            //create second row
            XWPFTableRow table4RowTwo = table4.createRow();

            XWPFRun run4t01 = table4RowTwo.getCell(0).addParagraph().createRun();
            run4t01.setFontFamily("Times New Roman"); run4t01.setFontSize(12); run4t01.setText(" Количество услуг");
            run4t01.setBold(true);
            table4RowTwo.getCell(0).removeParagraph(0);

            XWPFParagraph parag44 = myNewDoc.createParagraph();

            parag44.setAlignment(ParagraphAlignment.LEFT);
            parag44.setSpacingBetween(1.05);
            parag44.setSpacingAfterLines(0);
            XWPFRun run4401 = parag44.createRun();
            run4401.setFontFamily("Times New Roman");
            run4401.setFontSize(13);
            run4401.setBold(true);
            run4401.addBreak();
            run4401.setText("Заключение специалиста ");
            run4401.setText(" ________________________________________________________");
            run4401.addBreak();
            for (int i = 0; i < 9; i++) {
                run4401.setText("________________________________________________________________________________");
                run4401.addBreak();
            }
            run4401.addBreak();
            run4401.setText("Специалист по социальной работе  _________________  _____________________________");
            XWPFParagraph parag45 = myNewDoc.createParagraph();

            parag45.setAlignment(ParagraphAlignment.CENTER);
            parag45.setSpacingBetween(1.05);
            parag45.setSpacingAfterLines(0);
            XWPFRun run4501 = parag45.createRun();
            run4501.setFontFamily("Times New Roman");
            run4501.setFontSize(10);
            run4501.setBold(true);
            run4501.setText("                                           подпись                           расшифровка          ");

            //СТРАНИЦА 10

            XWPFParagraph parag46 = myNewDoc.createParagraph();

            parag46.setAlignment(ParagraphAlignment.LEFT);
            parag46.setSpacingBetween(1.05);
            parag46.setSpacingAfterLines(0);
            parag46.setPageBreak(true);

            XWPFRun run4601 = parag46.createRun();
            run4601.setFontFamily("Times New Roman");
            run4601.setFontSize(13);
            run4601.setBold(true);
            run4601.setText("4.3 СОЦИАЛЬНО-ПЕДАГОГИЧЕСКАЯ РЕАБИЛИТАЦИЯ");
            run4601.addBreak();

            XWPFRun run4602 = parag46.createRun();
            run4602.setFontFamily("Times New Roman");
            run4602.setFontSize(12);
            run4602.setBold(true);
            run4602.setText("Дата _____________________");
            run4602.addBreak();
            run4602.setText("4.3.1. Результаты социально-педагогической диагностики ");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("4.3.2. Программа социально-педагогической реабилитации: ");
            run4602.addBreak();
            run4602.setText("4.3.2.1. Социально-педагогическое консультирование ");
            run4602.setText("______________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("4.3.2.2. Коррекционное обучение: ");
            run4602.addBreak();
            run4602.setText("обучение социальному общению ");
            run4602.setText("_________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.setText("социальная независимость ");
            run4602.setText("______________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("персональная сохранность ");
            run4602.setText("______________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("компьютерная грамотность ");
            run4602.setText("_____________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("4.3.2.3. Педагогическая коррекция (логопед, дефектолог) ");
            run4602.setText("___________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("4.3.2.4. Педагогическое просвещение ");
            run4602.setText("_____________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("4.3.2.5. Социальные навыки ");
            run4602.setText("_____________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.setText("_______________________________________________________________________________________");
            run4602.addBreak();
            run4602.addBreak();
            run4602.setText("4.4. Итог проведенных мероприятий  по социально-педагогической реабилитации: ");

            //Таблица
            XWPFTable table5 = myNewDoc.createTable();
            table5.setWidth(10400);

            //create first row
            XWPFTableRow table5RowOne = table5.getRow(0);

            XWPFRun run5t1 =  table5RowOne.getCell(0).addParagraph().createRun();
            run5t1.setFontFamily("Times New Roman"); run5t1.setFontSize(10); run5t1.setText("Наименование услуги ");
            run5t1.setBold(true);
            table5RowOne.getCell(0).removeParagraph(0);

            XWPFRun run5t2 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t2.setFontFamily("Times New Roman"); run5t2.setFontSize(10); run5t2.setText("Социально-педагогическое консультирование ");
            run5t2.setBold(true);
            table5RowOne.getCell(1).removeParagraph(0);

            XWPFRun run5t3 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t3.setFontFamily("Times New Roman"); run5t3.setFontSize(10); run5t3.setText("Социально-педагогическая диагностика ");
            run5t3.setBold(true);
            table5RowOne.getCell(2).removeParagraph(0);

            XWPFRun run5t4 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t4.setFontFamily("Times New Roman"); run5t4.setFontSize(10); run5t4.setText("Коррекционное обучение ");
            run5t4.setBold(true);
            table5RowOne.getCell(3).removeParagraph(0);

            XWPFRun run5t5 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t5.setFontFamily("Times New Roman"); run5t5.setFontSize(10); run5t5.setText("Педагогическое просвещение ");
            run5t5.setBold(true);
            table5RowOne.getCell(4).removeParagraph(0);

            XWPFRun run5t6 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t6.setFontFamily("Times New Roman"); run5t6.setFontSize(10); run5t6.setText("Педагогическая коррекция ");
            run5t6.setBold(true);
            table5RowOne.getCell(5).removeParagraph(0);

            XWPFRun run5t7 = table5RowOne.addNewTableCell().addParagraph().createRun();
            run5t7.setFontFamily("Times New Roman"); run5t7.setFontSize(10); run5t7.setText("Обучение социальным навыкам ");
            run5t7.setBold(true);
            table5RowOne.getCell(6).removeParagraph(0);


            //create second row
            XWPFTableRow table5RowTwo = table5.createRow();

            XWPFRun run5t01 =  table5RowTwo.getCell(0).addParagraph().createRun();
            run5t01.setFontFamily("Times New Roman"); run5t01.setFontSize(10); run5t01.setText("Количество услуг ");
            run5t01.setBold(true);
            table5RowTwo.getCell(0).removeParagraph(0);


            XWPFParagraph parag47 = myNewDoc.createParagraph();

            parag47.setAlignment(ParagraphAlignment.LEFT);
            parag47.setSpacingBetween(1.05);
            parag47.setSpacingAfterLines(0);

            XWPFRun run4701 = parag47.createRun();
            run4701.setFontFamily("Times New Roman");
            run4701.setFontSize(13);
            run4701.setBold(true);
            run4701.addBreak();
            run4701.setText("Заключение социального педагога ");
            run4701.setText("________________________________________________");
            run4701.addBreak();
            run4701.setText("________________________________________________________________________________");
            run4701.addBreak();
            run4701.setText("________________________________________________________________________________");
            run4701.addBreak();
            run4701.setText("________________________________________________________________________________");
            run4701.addBreak();
            run4701.setText("________________________________________________________________________________");

            XWPFRun run4702 = parag47.createRun();
            run4702.setFontFamily("Times New Roman");
            run4702.setFontSize(12);
            run4702.setBold(true);
            run4702.addBreak();
            run4702.addBreak();
            run4702.setText("Социальный педагог _______________________    __________________________________________");

            XWPFParagraph parag48 = myNewDoc.createParagraph();

            parag48.setAlignment(ParagraphAlignment.CENTER);
            parag48.setSpacingBetween(1.05);
            parag48.setSpacingAfterLines(0);
            XWPFRun run4801 = parag48.createRun();
            run4801.setFontFamily("Times New Roman");
            run4801.setFontSize(10);
            run4801.setBold(true);
            run4801.setText("                       подпись                         расшифровка          ");

            // СТРАНИЦА 11

            XWPFParagraph parag49 = myNewDoc.createParagraph();

            parag49.setAlignment(ParagraphAlignment.CENTER);
            parag49.setSpacingBetween(1.00);
            parag49.setSpacingAfterLines(0);
            parag49.setPageBreak(true);

            XWPFRun run4901 = parag49.createRun();
            run4901.setFontFamily("Times New Roman");
            run4901.setFontSize(13);
            run4901.setBold(true);
            run4901.setText("VI. ПСИХОЛОГИЧЕСКАЯ РЕАБИЛИТАЦИЯ");
            run4901.addBreak();

            XWPFParagraph parag50 = myNewDoc.createParagraph();

            parag50.setAlignment(ParagraphAlignment.LEFT);
            parag50.setSpacingBetween(1.00);
            parag50.setSpacingAfterLines(0);

            XWPFRun run5001 = parag50.createRun();
            run5001.setFontFamily("Times New Roman");
            run5001.setFontSize(12);
            run5001.setBold(true);
            run5001.setText("6.1. Дата психологического исследования ______________________________");
            run5001.addBreak();
            run5001.setText("6.2. Жалобы:  __________________________________________________________________________");
            run5001.addBreak();
            run5001.setText("_______________________________________________________________________________________");
            run5001.addBreak();
            run5001.setText("_______________________________________________________________________________________");
            run5001.addBreak();

            XWPFRun run5002 = parag50.createRun();
            run5002.setFontFamily("Times New Roman");
            run5002.setFontSize(12);
            run5002.setBold(true);
            run5002.setText("6.3. Психологический статус:");
            run5002.addBreak();
            run5002.setUnderline(UnderlinePatterns.SINGLE);

            XWPFRun run5003 = parag50.createRun();
            run5003.setFontFamily("Times New Roman");
            run5003.setFontSize(12);
            run5003.setBold(true);
            run5003.setText("6.3.1. Психологический климат в семье:");

            XWPFRun run5004 = parag50.createRun();
            run5004.setFontFamily("Times New Roman");
            run5004.setFontSize(12);
            run5004.setText(" благоприятный/ неблагоприятный ");
            run5004.addBreak();
            run5004.setText("_______________________________________________________________________________________");
            run5004.addBreak();

            XWPFRun run5005 = parag50.createRun();
            run5005.setFontFamily("Times New Roman");
            run5005.setFontSize(12);
            run5005.setBold(true);
            run5005.setText("6.3.2. Социальная адаптированность: ");

            XWPFRun run5006 = parag50.createRun();
            run5006.setFontFamily("Times New Roman");
            run5006.setFontSize(12);
            run5006.setText("- высокая адаптированность");
            run5006.addBreak();
            run5006.setText("                                                                     - средняя адаптированность");
            run5006.addBreak();
            run5006.setText("                                                                     - низкая адаптированность");
            run5006.addBreak();
            run5006.setText("                                                                     - дезадаптированность");


            XWPFParagraph parag51 = myNewDoc.createParagraph();

            parag51.setAlignment(ParagraphAlignment.BOTH);
            parag51.setSpacingBetween(1.00);
            parag51.setSpacingAfterLines(0);

            XWPFRun run5101 = parag51.createRun();
            run5101.setFontFamily("Times New Roman");
            run5101.setFontSize(12);
            run5101.setBold(true);
            run5101.setText("6.3.3. Интересы, увлечения, хобби: ");

            XWPFRun run5102 = parag51.createRun();
            run5102.setFontFamily("Times New Roman");
            run5102.setFontSize(12);
            run5102.setText("коллекционирование/ отдых на природе/ туризм/ прогулки/ музыка/ пение/ танцы/ компьютерные игры/ интернет-технологии/ фотографии/ просмотр фильмов и телепередач/ чтение/ наука/ домашние животные/ рукоделие/ общение с друзьями, единомышленниками/ развивающие игры/ кулинария/ посещение культурных мероприятий/  занятия в кружках, клубах, студиях/ рыбалка/ цветоводство/ оздоровительные тренировки/ общественная деятельность/ рисование/  другое ___________________________________________ \t");
            run5102.addBreak();
            run5102.addBreak();

            XWPFRun run5103 = parag51.createRun();
            run5103.setFontFamily("Times New Roman");
            run5103.setFontSize(12);
            run5103.setBold(true);
            run5103.setText("6.3.4. Общая удовлетворенность жизнью: ");

            XWPFRun run5104 = parag51.createRun();
            run5104.setFontFamily("Times New Roman");
            run5104.setFontSize(12);
            run5104.setText("- полная удовлетворенность \t");
            run5104.addBreak();
            run5104.setText("                                                                            - частичная удовлетворенность\t");
            run5104.addBreak();
            run5104.setText("                                                                            - неудовлетворенность\t");
            run5104.addBreak();


            XWPFRun run5105 = parag51.createRun();
            run5105.setFontFamily("Times New Roman");
            run5105.setFontSize(12);
            run5105.setBold(true);
            run5105.setText("6.3.5. Отношение к болезни (классификация типов реакции на болезнь Личко, Иванова):");
            run5105.addBreak();

            XWPFRun run5106 = parag51.createRun();
            run5106.setFontFamily("Times New Roman");
            run5106.setFontSize(12);
            run5106.setText("гармонический тип/ эргопатический тип/ анизогнозический тип/ тревожный тип/ ипохондрический тип/ неврастенический тип/ меланхолический тип/ апатический тип/ сенситивный тип/ эгоцентрический тип/ паранойяльный тип/ дисфорический тип\t");
            run5106.addBreak();
            run5106.addBreak();

            XWPFRun run5107 = parag51.createRun();
            run5107.setFontFamily("Times New Roman");
            run5107.setFontSize(12);
            run5107.setBold(true);
            run5107.setText("6.3.6. Состояние  эмоционально-волевой  сферы \t");
            run5107.addBreak();
            run5107.setText("_______________________________________________________________________________________");
            run5107.addBreak();
            run5107.setText("_______________________________________________________________________________________");
            run5107.addBreak();
            run5107.setText("_______________________________________________________________________________________");
            run5107.addBreak();
            run5107.setText("_______________________________________________________________________________________");
            run5107.addBreak();
            run5107.setText("_______________________________________________________________________________________");

            // Таблица

            XWPFTable table6 = myNewDoc.createTable();
            table6.setWidth(10700);

            //create first row
            XWPFTableRow table6RowOne = table6.getRow(0);



            XWPFRun run6t1 =  table6RowOne.getCell(0).addParagraph().createRun();
            run6t1.setFontFamily("Times New Roman"); run6t1.setFontSize(11); run6t1.setText(" ");
            table6RowOne.getCell(0).removeParagraph(0);

            XWPFRun run6t2 = table6RowOne.addNewTableCell().addParagraph().createRun();
            run6t2.setFontFamily("Times New Roman"); run6t2.setFontSize(11); run6t2.setText(" Результаты  при поступлении ");

            table6RowOne.getCell(1).removeParagraph(0);

            XWPFRun run6t3 = table6RowOne.addNewTableCell().addParagraph().createRun();
            run6t3.setFontFamily("Times New Roman"); run6t3.setFontSize(11); run6t3.setText(" Результаты при выписке ");
            table6RowOne.getCell(2).removeParagraph(0);

            table6RowOne.getCell(0).setWidth("1000");
            table6RowOne.getCell(1).setWidth("3000");
            table6RowOne.getCell(2).setWidth("3000");

            //create second row
            XWPFTableRow table6RowTwo = table6.createRow();

            XWPFRun run6t01 = table6RowTwo.getCell(0).addParagraph().createRun();
            run6t01.setFontFamily("Times New Roman"); run6t01.setFontSize(11); run6t01.setText(" Самочувствие, активность, настроение");
            table6RowTwo.getCell(0).removeParagraph(0);

            table6RowTwo.setHeight(250);
            table6RowTwo.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFTableRow table6RowThree = table6.createRow();

            XWPFRun run6t02 = table6RowThree.getCell(0).addParagraph().createRun();
            run6t02.setFontFamily("Times New Roman"); run6t02.setFontSize(11); run6t02.setText(" Уровень тревоги");
            table6RowThree.getCell(0).removeParagraph(0);

            table6RowThree.setHeight(250);
            table6RowThree.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFTableRow table6RowFour = table6.createRow();

            XWPFRun run6t03 = table6RowFour.getCell(0).addParagraph().createRun();
            run6t03.setFontFamily("Times New Roman"); run6t03.setFontSize(11); run6t03.setText(" Уровень депрессии");
            table6RowFour.getCell(0).removeParagraph(0);

            table6RowFour.setHeight(250);
            table6RowFour.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag52 = myNewDoc.createParagraph();

            parag52.setAlignment(ParagraphAlignment.LEFT);
            parag52.setSpacingBetween(1.00);
            parag52.setSpacingAfterLines(0);

            XWPFRun run5201 = parag52.createRun();
            run5201.setFontFamily("Times New Roman");
            run5201.setFontSize(12);
            run5201.setBold(true);
            run5201.addBreak();
            run5201.setText("6.3.7. Уровень психологической адаптированности: ");

            XWPFRun run5202 = parag52.createRun();
            run5202.setFontFamily("Times New Roman");
            run5202.setFontSize(12);
            run5202.setText("оптимальный/ высокий/      ");
            run5202.addBreak();
            run5202.setText("                                                                                               низкий/ дезадаптационный ");
            run5202.addBreak();

            XWPFRun run5203 = parag52.createRun();
            run5203.setFontFamily("Times New Roman");
            run5203.setFontSize(12);
            run5203.setBold(true);
            run5203.setText("6.3.8. Личностно-характерологические особенности ");
            run5203.addBreak();

            XWPFRun run5204 = parag52.createRun();
            run5204.setFontFamily("Times New Roman");
            run5204.setFontSize(12);
            run5204.setText("_______________________________________________________________________________________");
            run5204.addBreak();
            run5204.setText("_______________________________________________________________________________________");
            run5204.addBreak();
            run5204.setText("_______________________________________________________________________________________");
            run5204.addBreak();
            run5204.setText("_______________________________________________________________________________________");
            run5204.addBreak();
            run5204.setText("_______________________________________________________________________________________");
            run5204.addBreak();
            run5204.setText("_______________________________________________________________________________________");


            // СТРАНИЦА 12

            XWPFParagraph parag53 = myNewDoc.createParagraph();

            parag53.setAlignment(ParagraphAlignment.LEFT);
            parag53.setSpacingBetween(1.00);
            parag53.setSpacingAfterLines(0);
            parag53.setPageBreak(true);

            XWPFRun run5301 = parag53.createRun();
            run5301.setFontFamily("Times New Roman");
            run5301.setFontSize(12);
            run5301.setBold(true);
            run5301.setText("6.3.9. Познавательные процессы: ");

            XWPFRun run5302 = parag53.createRun();
            run5302.setFontFamily("Times New Roman");
            run5302.setFontSize(12);
            run5302.setText("- в норме/  ");
            run5302.addBreak();
            run5302.setText("                                                             - наблюдаются нарушения: речи/ воображения/ мышления/");
            run5302.addBreak();
            run5302.setText("                                                               памяти/ внимания/ представления/ восприятия/ ощущения/ др.");
            run5302.addBreak();
            run5302.setText("                                                               _____________________________________________________");
            run5302.addBreak();

            XWPFRun run5303 = parag53.createRun();
            run5303.setFontFamily("Times New Roman");
            run5303.setFontSize(12);
            run5303.setBold(true);
            run5303.setUnderline(UnderlinePatterns.SINGLE);
            run5303.setText("6.4. Программа психологической помощи:");
            run5303.addBreak();

            XWPFRun run5304 = parag53.createRun();
            run5304.setFontFamily("Times New Roman");
            run5304.setFontSize(12);
            run5304.setBold(true);
            run5304.setText("6.4.1. Нуждается в психологической реабилитации: ");

            XWPFRun run5305 = parag53.createRun();
            run5305.setFontFamily("Times New Roman");
            run5305.setFontSize(12);
            run5305.setText("да/  нет");
            run5305.addBreak();
            run5305.addBreak();

            XWPFRun run5306 = parag53.createRun();
            run5306.setFontFamily("Times New Roman");
            run5306.setFontSize(12);
            run5306.setBold(true);
            run5306.setText("6.4.2. В каких видах психологической реабилитации нуждается: ");
            run5306.addBreak();

            XWPFRun run5307 = parag53.createRun();
            run5307.setFontFamily("Times New Roman");
            run5307.setFontSize(12);
            run5307.setText("                                                                                                - психологическое консультирование");
            run5307.addBreak();
            run5307.setText("                                                                                                - психологическая диагностика");
            run5307.addBreak();
            run5307.setText("                                                                                                - коррекция личностной сферы");
            run5307.addBreak();
            run5307.setText("                                                                                                - коррекция эмоционально-волевой сферы");
            run5307.addBreak();
            run5307.setText("                                                                                                - коррекция поведенческих аспектов");
            run5307.addBreak();
            run5307.setText("                                                                                                - коррекция межличностных отношений");
            run5307.addBreak();
            run5307.setText("                                                                                                - коррекция познавательной сферы");
            run5307.addBreak();
            run5307.setText("                                                                                                - психологическая разгрузка");
            run5307.addBreak();
            run5307.setText("                                                                                                - социально-психологический тренинг");
            run5307.addBreak();
            run5307.setText("                                                                                                - психопрофилактика");
            run5307.addBreak();
            run5307.addBreak();
            XWPFRun run5308 = parag53.createRun();
            run5308.setFontFamily("Times New Roman");
            run5308.setFontSize(12);
            run5308.setBold(true);
            run5308.setUnderline(UnderlinePatterns.SINGLE);
            run5308.setText("6.5. Цель и задачи реабилитации инвалида:");

            XWPFRun run5309 = parag53.createRun();
            run5309.setFontFamily("Times New Roman");
            run5309.setFontSize(12);
            run5309.setText(" улучшение психического здоровья инвалида");
            run5309.addBreak();
            run5309.addBreak();

            XWPFRun run5310 = parag53.createRun();
            run5310.setFontFamily("Times New Roman");
            run5310.setFontSize(12);
            run5310.setBold(true);
            run5310.setText("6.5.3. Реабилитационный потенциал: ");

            XWPFRun run5311 = parag53.createRun();
            run5311.setFontFamily("Times New Roman");
            run5311.setFontSize(12);
            run5311.setText("полностью сохранный/ относительно высокий");
            run5311.addBreak();
            run5311.setText("                                                                    удовлетворительный/ снижен/ значительно снижен");
            run5311.addBreak();

            XWPFRun run5312 = parag53.createRun();
            run5312.setFontFamily("Times New Roman");
            run5312.setFontSize(12);
            run5312.setBold(true);
            run5312.setText("6.5.4. Реабилитационный прогноз: ");

            XWPFRun run5313 = parag53.createRun();
            run5313.setFontFamily("Times New Roman");
            run5313.setFontSize(12);
            run5313.setText("положительный/ неопределенный/ отрицательный ");
            run5313.addBreak();
            run5313.addBreak();

            XWPFRun run5314 = parag53.createRun();
            run5314.setFontFamily("Times New Roman");
            run5314.setFontSize(12);
            run5314.setBold(true);
            run5314.setUnderline(UnderlinePatterns.SINGLE);
            run5314.setText("6.6. Итог мероприятий по психологической реабилитации:");

            // ТАБЛИЦА

            XWPFTable table7 = myNewDoc.createTable();
            table7.setWidth(10700);

            //create first row
            XWPFTableRow table7RowOne = table7.getRow(0);

            XWPFRun run7t1 = table7RowOne.getCell(0).addParagraph().createRun();
            run7t1.setFontFamily("Times New Roman");
            run7t1.setFontSize(10);
            run7t1.addBreak(); run7t1.addBreak(); run7t1.addBreak();
            run7t1.setText("Наименование услуги"); run7t1.addBreak();
//            table7RowOne.getCell(0).getCTTc().addNewTcPr().addNewTextDirection().setVal(STTextDirection.BT_LR);

            table7RowOne.getCell(0).removeParagraph(0);

            table7RowOne.setHeight(2100);
            table7RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            String s7 = "";

            for (int i = 1; i < 14; i++) {
                if (i == 1) {
                    s7 = "психолоогическое\n" +
                            "консультирование\n";
                } else if (i == 2) {
                    s7 = "психологическая \n" +
                            "диагностика\n";
                } else if (i == 3) {
                    s7 = "психологическая \n" +
                            "разгрузка\n";
                } else if (i == 4) {
                    s7 = "коррекция эмоционально-\n" +
                            "волевой сферы\n";
                } else if (i == 5) {
                    s7 = "коррекция\n" +
                            "познавательной сферы\n";
                } else if (i == 6) {
                    s7 = "коррекция межличностных \n" +
                            "отношений\n";
                } else if (i == 7) {
                    s7 = "коррекция \n" +
                            "поведенческих  аспектов\n";
                } else if (i == 8) {
                    s7 = "коррекция \n" +
                            "личностной сферы\n";
                } else if (i == 9) {
                    s7 = "социально-\n" +
                            "психологический тренинг\n";
                } else if (i == 10) {
                    s7 = "психопрофилактика";
                } else if (i == 11) {
                    s7 = "ЭЭГ-БОС";
                } else if (i == 12) {
                    s7 = "ЭМГ-БОС";
                } else if (i == 13) {
                    s7 = "ДАС-БОС";
                } else {
                    s7 = " ";
                }


                XWPFRun run7t2 = table7RowOne.addNewTableCell().addParagraph().createRun();
                run7t2.setFontFamily("Times New Roman");
                run7t2.setFontSize(10);
                run7t2.setText(s7);
                table7RowOne.getCell(i).removeParagraph(0);
                table7RowOne.getCell(i).getCTTc().addNewTcPr().addNewTextDirection().setVal(STTextDirection.BT_LR);

               table7RowOne.getCell(0).setWidth("900");
//                table7RowOne.getCell(1).setWidth("3000");
//                table7RowOne.getCell(2).setWidth("3000");
            }
            //create second row
            XWPFTableRow table7RowTwo = table7.createRow();

            XWPFRun run7t01 = table7RowTwo.getCell(0).addParagraph().createRun();
            run7t01.setFontFamily("Times New Roman"); run7t01.setFontSize(10); run7t01.setText("Количество");
            run7t01.addBreak();run7t01.setText("услуг");
            table7RowTwo.getCell(0).removeParagraph(0);

//            table7RowTwo.setHeight(250);
//            table7RowTwo.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag54 = myNewDoc.createParagraph();

            parag54.setAlignment(ParagraphAlignment.LEFT);
            parag54.setSpacingBetween(1.00);
            parag54.setSpacingAfterLines(0);

            XWPFRun run5401 = parag54.createRun();
            run5401.setFontFamily("Times New Roman");
            run5401.setFontSize(12);
            run5401.setBold(true);
            run5401.addBreak();
            run5401.setText("6.7. Оценка эффективности психокоррекционного воздействия");

            XWPFRun run5402 = parag54.createRun();
            run5402.setFontFamily("Times New Roman");
            run5402.setFontSize(12);
            for (int i = 0; i < 5; i++) {
                run5402.addBreak();
                run5402.setText("_______________________________________________________________________________________");

            }

            XWPFRun run5403 = parag54.createRun();
            run5403.setFontFamily("Times New Roman");
            run5403.setFontSize(12);
            run5403.setBold(true);
            run5403.addBreak();
            run5403.setText("6.8. Заключение психолога:");


            XWPFRun run5404 = parag54.createRun();
            run5404.setFontFamily("Times New Roman");
            run5404.setFontSize(12);
            for (int i = 0; i < 6; i++) {
                run5404.addBreak();
                run5404.setText("_______________________________________________________________________________________");

            }
            XWPFRun run5405 = parag54.createRun();
            run5405.setFontFamily("Times New Roman");
            run5405.setFontSize(12);
            run5405.setBold(true);
            run5405.addBreak();
            run5405.setText("                                                          Психолог _________________________________________________");

            // СТРАНИЦА 13

            XWPFParagraph parag55 = myNewDoc.createParagraph();

            parag55.setAlignment(ParagraphAlignment.CENTER);
            parag55.setSpacingBetween(1.00);
            parag55.setSpacingAfterLines(0);
            parag55.setPageBreak(true);

            XWPFRun run5501 = parag55.createRun();
            run5501.setFontFamily("Times New Roman");
            run5501.setFontSize(13);
            run5501.setBold(true);
            run5501.setText("VII. СОЦИОКУЛЬТУРНАЯ РЕАБИЛИТАЦИЯ");
            run5501.addBreak();

            XWPFParagraph parag56 = myNewDoc.createParagraph();

            parag56.setAlignment(ParagraphAlignment.BOTH);
            parag56.setSpacingBetween(1.00);
            parag56.setSpacingAfterLines(0);

            XWPFRun run5601 = parag56.createRun();
            run5601.setFontFamily("Times New Roman");
            run5601.setFontSize(12);
            run5601.setBold(true);
            run5601.setText("7.1. Дата проведения беседы: ____________________________________ \t");
            run5601.addBreak();
            run5601.addBreak();
            run5601.setText("7.2. Выявление сферы интересов: ");

            XWPFRun run5602 = parag56.createRun();
            run5602.setFontFamily("Times New Roman");
            run5602.setFontSize(12);
            run5602.setText("коллекционирование/ отдых на природе/ туризм/ прогулки/ музыка/ пение/ танцы/ компьютерные игры/ интернет-технологии/ фотографии/ просмотр фильмов и телепередач/ чтение/ наука/ домашние животные/ рукоделие/ общение с друзьями, единомышленниками/ развивающие игры/ кулинария/ посещение культурных мероприятий/ занятия в кружках, клубах, студиях/ рыбалка/ цветоводство/ оздоровительные тренировки/ общественная деятельность/ рисование/ другое __________________________________________________________");

            XWPFRun run5603 = parag56.createRun();
            run5603.setFontFamily("Times New Roman");
            run5603.setFontSize(12);
            run5603.setBold(true);
            run5603.addBreak();
            run5603.setText("7.3. Социокультурная реабилитация \t");
            run5603.addBreak();
            run5603.setText("7.3.1. Проведение мероприятий, направленных на создание условий возможности полноценного участия инвалидов в социокультурных мероприятиях, удовлетворяющих социокультурные и духовные запросы инвалидов, на расширение общего и культурного кругозора, сферы общения: ");
            XWPFRun run5604 = parag56.createRun();
            run5604.setFontFamily("Times New Roman");
            run5604.setFontSize(12);
            run5604.setText("посещение театров, выставок, экскурсии, встречи с деятелями литературы и искусства, праздники, юбилеи. \t");
            run5604.addBreak();
            run5604.addBreak();

            XWPFRun run5605 = parag56.createRun();
            run5605.setFontFamily("Times New Roman");
            run5605.setFontSize(12);
            run5605.setBold(true);
            run5605.setText("7.3.2. Содействие в обеспечении доступности для инвалидов посещений театров, музеев, кинотеатров, библиотек, возможности ознакомления с литературными произведениями и информацией о доступности учреждений культуры: ");
            XWPFRun run5606 = parag56.createRun();
            run5606.setFontFamily("Times New Roman");
            run5606.setFontSize(12);
            run5606.setText("подготовка и направление информационных материалов, обращений в органы и учреждения культуры, УСЗН, проведение круглых столов и других мероприятий для содействия в обеспечении доступности учреждений культуры. \t");
            run5606.addBreak();
            run5606.addBreak();

            XWPFRun run5607 = parag56.createRun();
            run5607.setFontFamily("Times New Roman");
            run5607.setFontSize(12);
            run5607.setBold(true);
            run5607.setText("7.3.3. Разработка и реализация разнопрофильных досуговых программ, способствующих формированию здоровой психики, развитию творческой инициативы и самостоятельности, направленных на обучение инвалидов навыкам проведения отдыха и досуга: ");
            XWPFRun run5608 = parag56.createRun();
            run5608.setFontFamily("Times New Roman");
            run5608.setFontSize(12);
            run5608.setText("общение, отдых, вечера встреч, прогулки, физкультурно-оздоровительная деятельность (игра в шашки, шахматы, дартс, теннис и др.), интеллектуально-познавательная деятельность активного (чтение, экскурсии, занятия в кружках, клубах и др.) и пассивного характера (просмотр телевизора, прослушивание музыки и др.), любительская деятельность прикладного характера (шитьё, фотодело, тестопластика, конструирование, моделирование и др.). \t");
            run5608.addBreak();
            run5608.addBreak();

            XWPFRun run5609 = parag56.createRun();
            run5609.setFontFamily("Times New Roman");
            run5609.setFontSize(12);
            run5609.setBold(true);
            run5609.setText("7.4. Итог проведенных мероприятий по социокультурной реабилитации: ");


            // Таблица

            XWPFTable table8 = myNewDoc.createTable();
            table8.setWidth(10700);

            //create first row
            XWPFTableRow table8RowOne = table8.getRow(0);



            XWPFRun run8t1 =  table8RowOne.getCell(0).addParagraph().createRun();
            run8t1.setFontFamily("Times New Roman"); run8t1.setFontSize(11); run8t1.setText(" Наименование");
            run8t1.addBreak(); run8t1.setText(" услуги");
            table8RowOne.getCell(0).removeParagraph(0);

            XWPFRun run8t2 = table8RowOne.addNewTableCell().addParagraph().createRun();
            run8t2.setFontFamily("Times New Roman"); run8t2.setFontSize(11); run8t2.setText(" Проведение мероприятий, направленных");
            run8t2.addBreak(); run8t2.setText(" на создание условий возможности ");
            run8t2.addBreak(); run8t2.setText(" полноценного участия инвалидов в ");
            run8t2.addBreak(); run8t2.setText(" социокультурных мероприятиях");
            table8RowOne.getCell(1).removeParagraph(0);

            XWPFRun run8t3 = table8RowOne.addNewTableCell().addParagraph().createRun();
            run8t3.setFontFamily("Times New Roman"); run8t3.setFontSize(11); run8t3.setText(" Оказание содействия");
            run8t3.addBreak(); run8t3.setText(" во взаимодействии с");
            run8t3.addBreak(); run8t3.setText(" учреждениями культуры");
            table8RowOne.getCell(2).removeParagraph(0);

            XWPFRun run8t4 = table8RowOne.addNewTableCell().addParagraph().createRun();
            run8t4.setFontFamily("Times New Roman"); run8t4.setFontSize(11); run8t4.setText(" Разработка и реализация");
            run8t4.addBreak(); run8t4.setText(" разнопрофильных ");
            run8t4.addBreak(); run8t4.setText(" досуговых программ");
            table8RowOne.getCell(3).removeParagraph(0);

            table8RowOne.getCell(0).setWidth("700");
            table8RowOne.getCell(1).setWidth("4000");
            table8RowOne.getCell(2).setWidth("2500");
            table8RowOne.getCell(3).setWidth("2500");
            //create second row
            XWPFTableRow table8RowTwo = table8.createRow();

            XWPFRun run8t01 = table8RowTwo.getCell(0).addParagraph().createRun();
            run8t01.setFontFamily("Times New Roman"); run8t01.setFontSize(11); run8t01.setText(" Количество услуг");
            table8RowTwo.getCell(0).removeParagraph(0);

//            table8RowTwo.setHeight(250);
//            table8RowTwo.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag57 = myNewDoc.createParagraph();

            parag57.setAlignment(ParagraphAlignment.LEFT);
            parag57.setSpacingBetween(1.00);
            parag57.setSpacingAfterLines(0);

            XWPFRun run5701 = parag57.createRun();
            run5701.setFontFamily("Times New Roman");
            run5701.setFontSize(12);
            run5701.setBold(true);
            run5701.setText("7.5. Заключение по итогам проведенной социокультурной реабилитации: ");
            for (int i = 0; i < 7; i++) {
                run5701.addBreak();
                run5701.setText("_______________________________________________________________________________________");

            }
            run5701.addBreak();
            run5701.setText("                                                           Культорганизатор ________________________________________");

// СТРАНИЦА 14

            XWPFParagraph parag58 = myNewDoc.createParagraph();

            parag58.setAlignment(ParagraphAlignment.CENTER);
            parag58.setSpacingBetween(1.05);
            parag58.setSpacingAfterLines(0);
            parag58.setPageBreak(true);

            XWPFRun run5801 = parag58.createRun();
            run5801.setFontFamily("Times New Roman");
            run5801.setFontSize(13);
            run5801.setBold(true);
            run5801.setText("VIII. Социально-оздоровительные мероприятия и спорт (адаптивная физическая культура)");


            XWPFParagraph parag59 = myNewDoc.createParagraph();
            parag59.setAlignment(ParagraphAlignment.BOTH);
            parag59.setSpacingBetween(1.25);
            parag59.setSpacingAfterLines(0);

            XWPFRun run5901 = parag59.createRun();
            run5901.setFontFamily("Times New Roman");
            run5901.setFontSize(11);
            run5901.addBreak();
            run5901.setText("8.1.Выполнение инвалидами под руководством персонала физических упражнений, в том числе аэробных, адекватных их физическим возможностям, оказывающих тренировочное действие и повышающих реабилитационные возможности, с проведением подбора, оптимизации физической нагрузки инвалидам, определением ее вида и объема ");

            XWPFRun run5902 = parag59.createRun();
            run5902.setFontFamily("Times New Roman");
            run5902.setFontSize(11);
            run5902.setBold(true);
            run5902.setText("(далее АФК)  _______________________________________________________ ");

            XWPFRun run5903 = parag59.createRun();
            run5903.setFontFamily("Times New Roman");
            run5903.setFontSize(11);
            run5903.addBreak();
            run5903.setText("_______________________________________________________________________________________________");
            run5903.addBreak();
            run5903.setText("_______________________________________________________________________________________________");
            run5903.addBreak();
            run5903.setText("_______________________________________________________________________________________________");
            run5903.addBreak();
            run5903.addBreak();
            run5903.setText("8.2.Содействие инвалидам в обеспечении доступности к объектам спортивно-оздоровительного назначения _______________________________________________________________________________________________");
            for (int i = 0; i < 5; i++) {
                run5903.addBreak();
                run5903.setText("_______________________________________________________________________________________________");

            }
            run5903.addBreak();
            run5903.addBreak();

            XWPFRun run5904 = parag59.createRun();
            run5904.setFontFamily("Times New Roman");
            run5904.setFontSize(12);
            run5904.setBold(true);
            run5904.setText("8.3. Итог проведенных мероприятий по адаптивной физкультуре");

            // Таблица

            XWPFTable table9 = myNewDoc.createTable();
            table9.setWidth(10700);

            //create first row
            XWPFTableRow table9RowOne = table9.getRow(0);



            XWPFRun run9t1 =  table9RowOne.getCell(0).addParagraph().createRun();
            run9t1.setFontFamily("Times New Roman"); run9t1.setFontSize(11); run9t1.setText(" Наименование");
            run9t1.addBreak(); run9t1.setText(" услуги"); run9t1.setBold(true);
            table9RowOne.getCell(0).removeParagraph(0);

            XWPFRun run9t2 = table9RowOne.addNewTableCell().addParagraph().createRun();
            run9t2.setFontFamily("Times New Roman"); run9t2.setFontSize(11); run9t2.setText(" Выполнение инвалидами под руководством персонала");
            run9t2.addBreak(); run9t2.setText(" физических упражнений, в том числе аэробных, ");
            run9t2.addBreak(); run9t2.setText(" адекватных их физическим возможностям, оказывающих");
            run9t2.addBreak(); run9t2.setText(" тренировочное действие и повышающих ");
            run9t2.addBreak(); run9t2.setText(" реабилитационные возможности, с проведением подбора");
            run9t2.addBreak(); run9t2.setText(" оптимизации физической нагрузки инвалидам, ");
            run9t2.addBreak(); run9t2.setText(" определением ее вида и объема");
            table9RowOne.getCell(1).removeParagraph(0);

            XWPFRun run9t3 = table9RowOne.addNewTableCell().addParagraph().createRun();
            run9t3.setFontFamily("Times New Roman"); run9t3.setFontSize(11); run9t3.setText(" содействие инвалидам в");
            run9t3.addBreak(); run9t3.setText(" обеспечении доступности");
            run9t3.addBreak(); run9t3.setText(" к объектам спортивно-");
            run9t3.addBreak(); run9t3.setText(" оздоровительного ");
            run9t3.addBreak(); run9t3.setText(" назначения");
            table9RowOne.getCell(2).removeParagraph(0);

            table9RowOne.getCell(0).setWidth("700");
            table9RowOne.getCell(1).setWidth("5500");
            table9RowOne.getCell(2).setWidth("2500");

            //create second row
            XWPFTableRow table9RowTwo = table9.createRow();

            XWPFRun run9t01 = table9RowTwo.getCell(0).addParagraph().createRun();
            run9t01.setFontFamily("Times New Roman"); run9t01.setFontSize(11); run9t01.setText(" Количество услуг");
            run9t01.setBold(true);
            table9RowTwo.getCell(0).removeParagraph(0);

//            table9RowOne.setHeight(1900);
//            table9RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag60 = myNewDoc.createParagraph();
            parag60.setAlignment(ParagraphAlignment.LEFT);
            parag60.setSpacingBetween(1.25);
            parag60.setSpacingAfterLines(0);

            XWPFRun run6001 = parag60.createRun();
            run6001.setFontFamily("Times New Roman");
            run6001.setFontSize(11);
            run6001.setBold(true);
            run6001.addBreak();
            run6001.setText("Заключение специалиста  _______________________________________________________________________ ");
            for (int i = 0; i < 8; i++) {
                run6001.addBreak();
                run6001.setText("_______________________________________________________________________________________________");

            }
            run6001.addBreak();
            run6001.addBreak();
            run6001.setText("Инструктор по адаптивной физкультуре     __________________       _______________________________");

            //СТРАНИЦА 15

            XWPFParagraph parag61 = myNewDoc.createParagraph();
            parag61.setAlignment(ParagraphAlignment.CENTER);
            parag61.setSpacingBetween(1.05);
            parag61.setSpacingAfterLines(0);
            parag61.setPageBreak(true);

            XWPFRun run6101 = parag61.createRun();
            run6101.setFontFamily("Times New Roman");
            run6101.setFontSize(13);
            run6101.setBold(true);
            run6101.setText("IX. Программа социальной реабилитации:");

            XWPFParagraph parag62 = myNewDoc.createParagraph();
            parag62.setAlignment(ParagraphAlignment.BOTH);
            parag62.setSpacingBetween(1.05);
            parag62.setSpacingAfterLines(0);

            XWPFRun run6201 = parag62.createRun();
            run6201.setFontFamily("Times New Roman");
            run6201.setFontSize(12);
            run6201.setText("7.1. Нуждается в социальной реабилитации:  да/ нет \t");
            run6201.addBreak();
            run6201.addBreak();
            run6201.setText("7.2. В каких видах социальной реабилитации нуждается: социально-средовая реабилитация, социально-бытовая адаптация, социально-педагогическая реабилитация, социокультурная реабилитация, психологическая реабилитация, социально-оздоровительные мероприятия и спорт");
            run6201.addBreak();
            run6201.addBreak();
            run6201.setText("7.3. ");
            XWPFRun run6202 = parag62.createRun();
            run6202.setFontFamily("Times New Roman");
            run6202.setFontSize(12);
            run6202.setBold(true);
            run6202.setText("Прогнозируемый результат: ");

            XWPFRun run6203 = parag62.createRun();
            run6203.setFontFamily("Times New Roman");
            run6203.setFontSize(12);
            run6203.setText("интеграция в общество (путем обеспечения необходимым набором ТСР, созданием доступной среды); \t");
            run6203.addBreak();
            run6203.setText("- формирование необходимых социально-бытовых навыков (путем обучения самообслуживанию и проведения мероприятий по обустройству жилища в соответствии с имеющимися ограничениями жизнедеятельности); \t");
            run6203.addBreak();
            run6203.setText("- коррекция и компенсация функций, приспособление к условиям социальной среды (педагогическими методами и средствами). \t");
            run6203.addBreak();
            run6203.setText("7.4. ");

            XWPFRun run6204 = parag62.createRun();
            run6204.setFontFamily("Times New Roman");
            run6204.setFontSize(12);
            run6204.setBold(true);
            run6204.setText("Заключение по итогам проведенных мероприятий по социальной реабилитации \t");
            run6204.addBreak();
            run6204.setText("_______________________________________________________________________________________");
            run6204.addBreak();
            run6204.addBreak();
            run6204.setText("Специалист по социальной работе ________________________________________________ \t");
            run6204.addBreak();
            run6204.setText("                                                                       (подпись) (расшифровка подписи, Ф.И.О) ");

// ТАБЛИЦА

            XWPFTable table10 = myNewDoc.createTable();

            XWPFTableRow table10RowOne = table10.getRow(0);


            XWPFRun run10t1 =  table10RowOne.getCell(0).addParagraph().createRun();
            run10t1.setFontFamily("Times New Roman"); run10t1.setFontSize(12); run10t1.setText(" № п/п ");
            table10RowOne.getCell(0).removeParagraph(0);

            String s10 = "";

            for (int i = 1; i <= 7; i++) {
                if (i == 1) {
                    s10 = " Дата";
                } else if (i == 2) {
                    s10 = " АД/утро";
                } else if (i == 3) {
                    s10 = " АД/вечер";
                } else if (i == 4) {
                    s10 = " PS утро";
                } else if (i == 5) {
                    s10 = " PS вечер";
                } else if (i == 6) {
                    s10 = " T утром";
                } else if (i == 7) {
                    s10 = " T вечер";
                } else s10 = " ";
                XWPFRun run10t2 = table10RowOne.addNewTableCell().addParagraph().createRun();
                run10t2.setFontFamily("Times New Roman");
                run10t2.setFontSize(12);
                run10t2.setText(s10);
                table10RowOne.getCell(i).removeParagraph(0);

            }
            for (int i = 0; i < 8; i++) {
                table10RowOne.getCell(i).setWidth("1500");
            }

//            table1RowOne.setHeight(300);
//            table1RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            for (int i = 1; i <= 29; i++) {
                String s1 = "";
                if (i == 0) {
                    s1 = "1. Диета";
                } else
                if (i == 1) {
                    s1 = "2. ЛФК";
                } else
                if (i == 2) {
                    s1 = "3. Консультация физиотерапевта";
                } else {
                    s1 = " ";
                }
                XWPFTableRow table10RowTwo = table10.createRow();
                table10RowTwo.setHeight(250);
                table10RowTwo.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

                XWPFRun run10t01 = table10RowTwo.getCell(0).addParagraph().createRun();
                run10t01.setFontFamily("Times New Roman");
                run10t01.setFontSize(12);
                run10t01.setText(" " + String.valueOf(i));
                table10RowTwo.getCell(0).removeParagraph(0);


            }

            // СТРАНИЦА 16

            XWPFParagraph parag63 = myNewDoc.createParagraph();
            parag63.setAlignment(ParagraphAlignment.CENTER);
            parag63.setSpacingBetween(1.05);
            parag63.setSpacingAfterLines(0);
            parag63.setPageBreak(true);

            XWPFRun run6301 = parag63.createRun();
            run6301.setFontFamily("Times New Roman");
            run6301.setFontSize(13);
            run6301.setBold(true);
            run6301.setText("VIII. ПРЕДОСТАВЛЕНИЕ СОЦИАЛЬНЫХ УСЛУГ");

            XWPFParagraph parag64 = myNewDoc.createParagraph();
            parag64.setAlignment(ParagraphAlignment.LEFT);
            parag64.setSpacingBetween(1.15);
            parag64.setSpacingAfterLines(0);

            XWPFRun run6401 = parag64.createRun();
            run6401.setFontFamily("Times New Roman");
            run6401.setFontSize(12);
            run6401.setBold(true);
            run6401.setText("8.1. Предоставление социально-психологических услуг:");
            run6401.addBreak();
            run6401.addBreak();

// Таблица

            XWPFTable table11 = myNewDoc.createTable();
            table11.setWidth(10700);

            //create first row
            XWPFTableRow table11RowOne = table11.getRow(0);



            XWPFRun run11t1 =  table11RowOne.getCell(0).addParagraph().createRun();
            run11t1.setFontFamily("Times New Roman"); run11t1.setFontSize(11); run11t1.setText(" Наименование");
            run11t1.addBreak(); run11t1.setText(" услуги");
            table11RowOne.getCell(0).removeParagraph(0);

            XWPFRun run11t2 = table11RowOne.addNewTableCell().addParagraph().createRun();
            run11t2.setFontFamily("Times New Roman"); run11t2.setFontSize(11); run11t2.setText(" проведение мероприятий по психологической");
            run11t2.addBreak(); run11t2.setText(" разгрузке инвалидов, с использованием сенсорного");
            run11t2.addBreak(); run11t2.setText(" оборудования, приборов для аромотерапии,");
            run11t2.addBreak(); run11t2.setText(" аудио-видеоаппаратуры");
            table11RowOne.getCell(1).removeParagraph(0);

            XWPFRun run11t3 = table11RowOne.addNewTableCell().addParagraph().createRun();
            run11t3.setFontFamily("Times New Roman"); run11t3.setFontSize(11); run11t3.setText(" проведение занятий в группах");
            run11t3.addBreak(); run11t3.setText(" взаимоподдержки, клубах общения,");
            run11t3.addBreak(); run11t3.setText(" психопрофилактики");
            table11RowOne.getCell(2).removeParagraph(0);

            table11RowOne.getCell(0).setWidth("700");
            table11RowOne.getCell(1).setWidth("4700");
            table11RowOne.getCell(2).setWidth("3300");

            //create second row
            XWPFTableRow table11RowTwo = table11.createRow();

            XWPFRun run11t01 = table11RowTwo.getCell(0).addParagraph().createRun();
            run11t01.setFontFamily("Times New Roman"); run11t01.setFontSize(11); run11t01.setText(" Количество услуг");
            table11RowTwo.getCell(0).removeParagraph(0);

//            table9RowOne.setHeight(1900);
//            table9RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag65 = myNewDoc.createParagraph();
            parag65.setAlignment(ParagraphAlignment.LEFT);
            parag65.setSpacingBetween(1.15);
            parag65.setSpacingAfterLines(0);

            XWPFRun run6501 = parag65.createRun();
            run6501.setFontFamily("Times New Roman");
            run6501.setFontSize(12);
            run6501.setBold(true);
            run6501.addBreak();
            run6501.setText("Психолог  ____________________     ______________________________");
            run6501.addBreak();

            XWPFRun run6502 = parag65.createRun();
            run6502.setFontFamily("Times New Roman");
            run6502.setFontSize(12);
            run6502.setText("                                       (подпись)             (расшифровка подписи, Ф.И.О.)  ");
            run6502.addBreak();



            XWPFParagraph parag66 = myNewDoc.createParagraph();
            parag66.setAlignment(ParagraphAlignment.LEFT);
            parag66.setSpacingBetween(1.15);
            parag66.setSpacingAfterLines(0);

            XWPFRun run6601 = parag66.createRun();
            run6601.setFontFamily("Times New Roman");
            run6601.setFontSize(12);
            run6601.setBold(true);
            run6601.setText("8.2. Предоставление социально-педагогических услуг:");
            run6601.addBreak();

// Таблица

            XWPFTable table12 = myNewDoc.createTable();
            table12.setWidth(10700);

            //create first row
            XWPFTableRow table12RowOne = table12.getRow(0);



            XWPFRun run12t1 =  table12RowOne.getCell(0).addParagraph().createRun();
            run12t1.setFontFamily("Times New Roman"); run12t1.setFontSize(12); run12t1.setText(" Наименование");
            run12t1.addBreak(); run12t1.setText(" услуги");
            table12RowOne.getCell(0).removeParagraph(0);

            XWPFRun run12t2 = table12RowOne.addNewTableCell().addParagraph().createRun();
            run12t2.setFontFamily("Times New Roman"); run12t2.setFontSize(12); run12t2.setText(" Проведение занятий в ");
            run12t2.addBreak(); run12t2.setText(" группах взаимопомощи,");
            run12t2.addBreak(); run12t2.setText(" клубах общения");
            table12RowOne.getCell(1).removeParagraph(0);

            XWPFRun run12t3 = table12RowOne.addNewTableCell().addParagraph().createRun();
            run12t3.setFontFamily("Times New Roman"); run12t3.setFontSize(12); run12t3.setText(" Оказание содействия в профессиональной ");
            run12t3.addBreak(); run12t3.setText(" ориентации в части профессионального");
            run12t3.addBreak(); run12t3.setText(" консультирования и информирования");
            table12RowOne.getCell(2).removeParagraph(0);

            table12RowOne.getCell(0).setWidth("700");
            table12RowOne.getCell(1).setWidth("3000");
            table12RowOne.getCell(2).setWidth("5000");

            //create second row
            XWPFTableRow table12RowTwo = table12.createRow();

            XWPFRun run12t01 = table12RowTwo.getCell(0).addParagraph().createRun();
            run12t01.setFontFamily("Times New Roman"); run12t01.setFontSize(12); run12t01.setText(" Количество услуг");
            table12RowTwo.getCell(0).removeParagraph(0);

//            table9RowOne.setHeight(1900);
//            table9RowOne.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);

            XWPFParagraph parag67 = myNewDoc.createParagraph();
            parag67.setAlignment(ParagraphAlignment.LEFT);
            parag67.setSpacingBetween(1.15);
            parag67.setSpacingAfterLines(0);

            XWPFRun run6701 = parag67.createRun();
            run6701.setFontFamily("Times New Roman");
            run6701.setFontSize(12);
            run6701.setBold(true);
            run6701.addBreak();
            run6701.setText("Специалист по социальной работе  ____________________     ______________________________");
            run6701.addBreak();

            XWPFRun run6702 = parag67.createRun();
            run6702.setFontFamily("Times New Roman");
            run6702.setFontSize(12);
            run6702.setText("                                                                                   (подпись)             (расшифровка подписи, Ф.И.О.)  ");
            run6702.addBreak();



            myNewDoc.write(fos);
            fos.close();

            System.out.println("Document created");
        } catch (Exception e) {
            System.out.println("Something is wrong");
        }

    }

}
// XWPFTable table = myNewDoc.createTable();
//
//            //create first row
//            XWPFTableRow tableRowOne = table.getRow(0);


//            tableRowOne.getCell(0).setText("col one, row one");
//            tableRowOne.addNewTableCell().setText("col two, row one");
//            tableRowOne.addNewTableCell().setText("col three, row one");
//            tableRowOne.addNewTableCell().setText("col four, row one");
//
//            //create second row
//            XWPFTableRow tableRowTwo = table.createRow();
//            tableRowTwo.getCell(0).setText("col one, row two");
//            tableRowTwo.getCell(1).setText("col two, row two");
//            tableRowTwo.getCell(2).setText("col three, row two");
//            tableRowTwo.getCell(3).setText("col four, row two");
//
//            //create third row
//            XWPFTableRow tableRowThree = table.createRow();
//            tableRowThree.getCell(0).setText("col one, row three");
//            tableRowThree.getCell(1).setText("col two, row three");
//            tableRowThree.getCell(2).setText("col three, row three");
//            tableRowThree.getCell(3).setText("col four, row three");
//
//            XWPFTableRow tableRowFour = table.createRow();
//            tableRowFour.getCell(0).setText("col one, row four");
//            tableRowFour.getCell(1).setText("col two, row four");
//            tableRowFour.getCell(2).setText("col three, row four");
//            tableRowFour.getCell(3).setText("col four, row four");
//
//            XWPFTableRow tableRowFive = table.createRow();
//            tableRowFive.getCell(0).setText("col one, row five");
//            tableRowFive.getCell(1).setText("col two, row five");
//            tableRowFive.getCell(2).setText("col three, row five");
//            tableRowFive.getCell(3).setText("col four, row five");
//
//            XWPFTableRow tableRowSix = table.createRow();
//            tableRowSix.getCell(0).setText("col one, row six");
//            tableRowSix.getCell(1).setText("col two, row six");
//            tableRowSix.getCell(2).setText("col three, row six");
//            tableRowSix.getCell(3).setText("col four, row six");



// parag.setIndentationHanging(1000); отступ слева
// parag.setSpacingAfter(1000); отступ снизу
//  parag.setBorderBottom(Borders.BASIC_THIN_LINES); линии снизу
// parag.setBorderTop(Borders.BASIC_THIN_LINES); линии сверх параграфа
// parag.setPageBreak(true); разрыв страницы (с новой страницы)
// parag.setNumID(BigInteger.ONE); нумерация
// run1.setStrike(true); зачеркивание
// run1.setUnderline(UnderlinePatterns.SINGLE); подчеркивание
// run.setTextPosition(100);
