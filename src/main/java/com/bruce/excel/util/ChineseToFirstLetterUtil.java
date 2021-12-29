package com.bruce.excel.util;

import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinCaseType;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;
import net.sourceforge.pinyin4j.format.exception.BadHanyuPinyinOutputFormatCombination;
import org.apache.commons.lang3.StringUtils;

import java.text.MessageFormat;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/12/28 23:45
 * @Author Bruce
 */
public class ChineseToFirstLetterUtil {

    /**
     * 获取拼音首字母
     * @param str
     * @return
     */
    public static String chineseToFirstLetter(String str) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < str.length(); i++) {
            String character = String.valueOf(str.charAt(i));
            String pinying = converterToFirstSpell(character);
            if (StringUtils.isNotEmpty(pinying)) {
                sb.append(pinying.charAt(0));
            }
        }
        return sb.toString();
    }

    /**
     * 获取中文
     * @param str
     * @return
     */
    public static String chineseLetter(String str) {
        String reg = "[^\u4e00-\u9fa5]";
        str = str.replaceAll(reg, "");
        return str;
    }

    private static String converterToFirstSpell(String chinese) {
        StringBuilder pinyinName = new StringBuilder();
        char[] nameChar = chinese.toCharArray();
        HanyuPinyinOutputFormat defaultFormat = new HanyuPinyinOutputFormat();
        defaultFormat.setCaseType(HanyuPinyinCaseType.LOWERCASE);
        defaultFormat.setToneType(HanyuPinyinToneType.WITHOUT_TONE);
        for (char c : nameChar) {
            String s = String.valueOf(c);
            if (s.matches("[\\u4e00-\\u9fa5]")) {
                try {
                    String[] mPinyinArray = PinyinHelper.toHanyuPinyinStringArray(c, defaultFormat);
                    pinyinName.append(mPinyinArray[0]);
                } catch (BadHanyuPinyinOutputFormatCombination e) {
                    e.printStackTrace();
                }
            }
        }
        return pinyinName.toString();
    }


    public static void main(String[] args) {
        String str = "我爱 中华!（zzz_ss）";
        System.out.println(chineseToFirstLetter(str));
        System.out.println(chineseLetter(str));
        System.out.println(MessageFormat.format("{0}==>''{1}''", chineseToFirstLetter(str), chineseLetter(str)));
    }
}
