# poi
    word
        XWPFDocument 创建的MS-Word文档与.docx文件格式。





            设置段落背景颜色
            CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
            cTShd.setVal(STShd.CLEAR);
            cTShd.setFill("97FFFF");

            表格不需要边框     XWPFTable infoTable
            infoTable.getCTTbl().getTblPr().unsetTblBorders();

            建立一个表格的时候设置列宽跟随内容伸缩
            CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
            infoTableWidth.setType(STTblWidth.DXA);
            infoTableWidth.setW(BigInteger.valueOf(9072));

            换行符号

            　　硬换行：文件中换行，如果是键盘中使用了"enter"的换行。

            　　软换行：文件中一行的字符数容量有限，当字符数量超过一定值时，会自动切到下行显示。

            　　对程序来说，硬换行才是可以识别的、确定的换行，软换行与字体大小、缩进有关。



