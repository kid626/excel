package com.bruce.excel.service;

import com.bruce.excel.entity.IndexOrNameData;

import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
public interface IndexOrNameDataService {

    /**
     * 保存
     *
     * @param list List<IndexOrNameData>
     */
    void save(List<IndexOrNameData> list);

}
