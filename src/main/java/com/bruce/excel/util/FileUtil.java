package com.bruce.excel.util;


import org.apache.commons.io.FileUtils;
import org.apache.poi.util.IOUtils;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
public class FileUtil {

    /**
     * MultipartFile 转 File
     *
     * @param multipartFile MultipartFile
     * @return File
     * @throws IOException multipartFile.getInputStream()
     */
    public static File multipartFileToFile(MultipartFile multipartFile) throws IOException {
        String path = "./temp";
        File file = new File(path);
        FileUtils.copyInputStreamToFile(multipartFile.getInputStream(), file);
        return file;
    }

    /**
     * File 转 MultipartFile
     *
     * @param file File
     * @return MultipartFile
     * @throws Exception new FileInputStream(file)
     */
    public static MultipartFile fileToMultipartFile(File file) throws Exception {
        FileInputStream input = new FileInputStream(file);
        return new MockMultipartFile("file", file.getName(), "text/plain", IOUtils.toByteArray(input));
    }


}
