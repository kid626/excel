package com.bruce.excel.service;

import com.bruce.excel.entity.Test;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
public interface TestService {

    /**
     * 批量导入
     *
     * @param file     文件
     * @param response HttpServletResponse
     * @return List<Test>
     */
    List<Test> batchSave(MultipartFile file, HttpServletResponse response) throws Exception;


    void save(List<Test> data);

}
