package cn.xxx.xmind2excel.biz;

import lombok.Data;

/**
 * @author xiongchenghui
 * @date 2020-11-09
 * &Desc 测试用例详情
 */
@Data
public class TestCaseInfo {
    /** 所有用例总数 */
    private Integer testCaseNo;

    /** 所有用例步骤数 */
    private Integer testCaseSteps;

    /** 所有用例检查点数 */
    private Integer testCaseCheckPointers;
}