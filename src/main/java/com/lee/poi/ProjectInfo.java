package com.lee.poi;

import cn.hutool.core.lang.Validator;
import cn.hutool.core.util.ReUtil;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * 作者 yhl
 * 日期 2020-01-05 14:02:19
 * 描述
 */
@Data
@Getter
@Setter
public class ProjectInfo implements Serializable {
    /**
     * 用户id
     */
    protected Integer projectId;

    /**
     * 项目/成果名称
     */
    protected String projectName;

    /**
     * 应用领域
     */
    protected List<Integer> useArea;

    /**
     * 技能领域
     */
    protected List<Integer> skillArea;

    /**
     * 估值：单位-万元
     */
    protected BigDecimal valuation;

    /**
     * 估值方式：1 面议，2 已估
     */
    protected Integer valuationType;

    /**
     * 成果类型
     */
    protected List<Integer> fruitType;

    /**
     * 当前阶段
     */
    protected List<Integer> processType;

    /**
     * 融资阶段
     */
    protected List<Integer> financingType;

    /**
     * 合作形式：多选
     */
    protected List<String> cooperateType;

    /**
     * 您的诉求：多选
     */
    protected List<String> yourWant;

    /**
     * 您的诉求描述
     */
    protected String yourWantContent;

    /**
     * 关键词
     */
    protected List<String> words;

    /**
     * 单位名称
     */
    protected String companyName;

    /**
     * 联系人姓名
     */
    protected String contactName;

    /**
     * 手机号
     */
    protected String phoneNumber;

    /**
     * 省id
     */
    protected Integer provinceId;

    /**
     * 城市id
     */
    protected Integer cityId;

    /**
     * 区域id
     */
    protected Integer areaId;

    /**
     * 宣传图片
     */
    protected List<String> photos;

    private String summary;

    /**
     * 项目详情
     */
    protected String detail;

    /**
     * 证明材料，多个url
     */
    protected List<String> proveUrl;

    /**
     * 创建时间
     */
    protected Date createTime;

    /**
     * 最后更新时间
     */
    protected Date updateTime;

    /**
     * 备注
     */
    protected String comment;

    /**
     * 状态
     */
    protected Integer state;

    protected Integer approvalState;

    protected String approvalOpinion;

    protected String approvalUser;

    protected Date approvalTime;

    protected Integer delFlag;

    protected String startVideo;

    protected String endVideo;

    protected Integer browseNum;

    private String checkSolution;

    private String indexExcelLabel;

    private String useAreaExcelLabel;

    private String skillAreaExcelLabel;

    private String valuationExcelLabel;

    private String fruitTypeExcelLabel;

    private String yourWantExcelLabel;

    private String processTypeExcelLabel;

    private String financingTypeExcelLabel;

    private String cooperateTypeExcelLabel;

    private String provinceExcelLabel;

    private String wordsExcelLabel;

    private String provinceName;

    private String cityName;

    private Short importFlag;

    //excel导入转换项目方联系人姓名
    public void convertsetContactName(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 项目方联系人姓名为空");
        else
            this.contactName = str;
    }

    //excel导入转换项目方手机号
    public void convertsetPhoneNumber(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 项目方手机号为空");
        else if (!Validator.isMobile(str))
            throw new Exception("编号 " + indexExcelLabel + " 项目方手机号格式错误");
        else
            this.phoneNumber = str;
    }

    //excel导入转换项目名称
    public void convertsetProjectName(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 项目名称为空");
        else {
            if (str.length() > 30)
                this.projectName = str.substring(0, 30);
            else
                this.projectName = str;
        }
    }

    //excel导入转换诉求描述
    public void convertsetYourWantContent(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 诉求描述为空");
        else if (str.length() > 150) {
            this.yourWantContent = str.substring(0, 150);
        } else {
            this.yourWantContent = str;
        }
    }

    //excel导入转换公司名称
    public void convertsetCompanyName(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 公司名称为空");
        else
            this.companyName = str;
    }

    //excel导入转换项目简介
    public void convertsetSummary(String str) throws Exception {
        if (StringUtils.isBlank(str))
            throw new Exception("编号 " + indexExcelLabel + " 项目简介为空");
        else {
            if (str.length() > 100)
                this.summary = str.substring(0, 100);
            else
                this.summary = str;
        }
    }

    //excel导入转换项目详情
    public void convertsetDetail(String str) throws Exception {
        if (StringUtils.isBlank(str))
            return;
        else {
            if (str.length() > 2000)
                this.detail = str.substring(0, 2000);
            else
                this.detail = str;
        }
    }

    //excel导入转换应用领域
    public void convertsetUseAreaExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.useArea = new ArrayList<Integer>();
                this.useArea.add(ReUtil.getFirstNumber(str));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 应用领域格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 应用领域为空");
        }
    }

    //excel导入转换技术领域
    public void convertsetSkillAreaExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.skillArea = new ArrayList<Integer>();
                this.skillArea.add(ReUtil.getFirstNumber(str));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 技术领域格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 技术领域为空");
        }
    }

    //excel导入转换成果类型
    public void convertsetFruitTypeExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.fruitType = new ArrayList<Integer>();
                this.fruitType.add(ReUtil.getFirstNumber(str));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 成果类型格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 成果类型为空");
        }
    }

    //excel导入转换项目诉求
    public void convertsetYourWantExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.yourWant = new ArrayList<String>();
                this.yourWant = Arrays.asList(str.replaceAll("，", ",").split(","));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 项目诉求格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 项目诉求为空");
        }
    }

    //excel导入转换当前阶段
    public void convertsetProcessTypeExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.processType = new ArrayList<Integer>();
                this.processType.add(ReUtil.getFirstNumber(str));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 当前阶段格式错误");
            }
        }
    }

    //excel导入转换融资类型
    public void convertsetFinancingTypeExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.financingType = new ArrayList<Integer>();
                this.financingType.add(ReUtil.getFirstNumber(str));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 融资类型格式错误");
            }
        } /*else {
            throw new Exception("编号 " + indexExcelLabel + " 融资类型为空");
        }*/
    }

    //excel导入转换归属方性质
    public void convertsetCooperateTypeExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.cooperateType = Arrays.asList(str.replaceAll("，", ",").split(","));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 归属方性质格式错误");
            }
        } /*else {
            throw new Exception("编号 " + indexExcelLabel + " 归属方性质为空");
        }*/
    }

    //excel导入转换估值
    public void convertsetValuationExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                if ("面议".equalsIgnoreCase(str)) {
                    valuationType = 2;
                    valuation = new BigDecimal(0.0);
                } else {
                    valuation = new BigDecimal(str);
                    valuationType = 1;
                }
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 估值格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 估值为空");
        }
    }

    //excel导入转换省市名称
    public void convertsetProvinceExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.provinceName = str.split("/")[0].replace("省", "");
                this.cityName = str.split("/")[1].replace("市", "");
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 项目方所在省市格式错误");
            }
        } else {
            throw new Exception("编号 " + indexExcelLabel + " 项目方所在省市为空");
        }
    }

    //excel导入转换搜索关键字
    public void convertsetWordsExcelLabel(String str) throws Exception {
        if (StringUtils.isNotBlank(str)) {
            try {
                this.words = Arrays.asList(str.replaceAll("，", ",").split(","));
            } catch (Exception e) {
                throw new Exception("编号 " + indexExcelLabel + " 关键字格式错误");
            }
        } else
            throw new Exception("编号 " + indexExcelLabel + " 关键字为空");
    }

}