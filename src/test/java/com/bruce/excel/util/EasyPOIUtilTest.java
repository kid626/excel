package com.bruce.excel.util;


import com.bruce.excel.entity.FileEntity;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Slf4j
public class EasyPOIUtilTest {

    @Test
    public void testEasyImport() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("easypoi/in/easy_poi_import_test.xlsx");
        File file = classPathResource.getFile();
        // File 转 MultipartFile
        FileInputStream input = new FileInputStream(file);
        MultipartFile multipartFile = new MockMultipartFile("file", file.getName(), "text/plain", IOUtils.toByteArray(input));
        List<FileEntity> list = EasyPOIUtil.importExcel(multipartFile, 1, 1, FileEntity.class);
        for (FileEntity fileEntity : list) {
            System.out.println(fileEntity);
        }

    }

    @Test
    public void testEasyExport() throws Exception {
        List<FileEntity> list = new ArrayList<>();
        FileEntity fileEntity = FileEntity.builder().name("testName").gender("2").date(new Date()).tags("testTags").build();
        list.add(fileEntity);
        FileEntity fileEntity1 = FileEntity.builder().name("testName1").gender("1").date(new Date()).tags("testTags1").build();
        list.add(fileEntity1);
        EasyPOIUtil.exportExcel(list, "testTitle", "testSheet", FileEntity.class, "src/main/resources/easypoi/out/easy_poi_export_test.xlsx", null);
    }


}