# app.py
import streamlit as st
import pandas as pd
import os, re, tempfile, shutil
from pathlib import Path


FIELD_MAPPING = {
    "通讯地址": ("Address", "字符型", 300, 0),
    "法人代表身份证件代码": ("InstReprIDCode", "字符型", 40, 0),
    "法人代表证件类型": ("InstReprIDType", "字符型", 3, 0),
    "法人代表姓名": ("InstReprName", "字符型", 60, 0),
    "申请单编号": ("AppSheetSerialNo", "字符型", 24, 0),
    "个人证件类型及机构证件型": ("CertificateType", "字符型", 3, 0),
    "投资人证件号码": ("CertificateNo", "字符型", 40, 0),
    "投资人户名": ("InvestorName", "字符型", 200, 0),
    "交易发生日期": ("TransactionDate", "数字字符型,限于0-9", 8, 0),
    "交易发生时间": ("TransactionTime", "数字字符型,限于0-9", 6, 0),
    "个人/机构标志": ("IndividualOrInstitution", "数字字符型,限于0-9", 1, 0),
    "投资人邮政编码": ("PostCode", "数字字符型,限于0-9", 6, 0),
    "经办人证件号码": ("TransactorCertNo", "字符型", 40, 0),
    "经办人证件类型": ("TransactorCertType", "字符型", 3, 0),
    "经办人姓名": ("TransactorName", "字符型", 60, 0),
    "投资人基金交易帐号": ("TransactionAccountID", "字符型", 17, 0),
    "销售人代码": ("DistributorCode", "字符型", 9, 0),
    "业务代码": ("BusinessCode", "字符型", 3, 0),
    "基金管理人在资金清算机构的交收帐号": ("AcctNoOfFMInClearingAgency", "字符型", 28, 0),
    "基金管理人在资金清算机构的交收账户名": ("AcctNameOfFMInClearingAgency", "字符型", 60, 0),
    "基金资金清算机构代码": ("ClearingAgencyCode", "数字字符型,限于0-9", 9, 0),
    "投资人出生日期": ("InvestorsBirthday", "数字字符型,限于0-9", 8, 0),
    "投资人在销售人处用于交易的资金帐号": ("DepositAcct", "字符型", 40, 0),
    "交易所在地区编号": ("RegionCode", "数字字符型,限于0-9", 4, 0),
    "投资人学历": ("EducationLevel", "字符型", 3, 0),
    "投资人E-MAIL地址": ("EmailAddress", "字符型", 40, 0),
    "投资人传真号码": ("FaxNo", "字符型", 40, 0),
    "投资人职业代码": ("VocationCode", "字符型", 5, 0),
    "投资人住址电话": ("HomeTelNo", "字符型", 40, 0),
    "投资人年收入": ("AnnualIncome", "数值型，其长度不包含小数点，可参与数值计算", 16, 0),
    "投资人手机号码": ("MobileTelNo", "字符型", 40, 0),
    "网点号码": ("BranchCode", "字符型", 9, 0),
    "投资人单位电话号码": ("OfficeTelNo", "字符型", 40, 0),
    "投资人户名简称": ("AccountAbbr", "字符型", 20, 0),
    "密函编号": ("ConfidentialDocumentCode", "字符型", 8, 0),
    "投资人性别": ("Sex", "数字字符型,限于0-9", 1, 0),
    "上海证券帐号": ("SHSecuritiesAccountID", "字符型", 10, 0),
    "深圳证券帐号": ("SZSecuritiesAccountID", "字符型", 10, 0),
    "投资人基金帐号": ("TAAccountID", "字符型", 12, 0),
    "投资人电话号码": ("TelNo", "字符型", 40, 0),
    "使用的交易手段": ("TradingMethod", "字符型", 8, 0),
    "未成年人标志": ("MinorFlag", "字符型", 1, 0),
    "对帐单寄送选择": ("DeliverType", "字符型", 1, 0),
    "经办人识别方式": ("TransactorIDType", "字符型", 1, 0),
    "基金账户卡的凭证号": ("AccountCardID", "字符型", 8, 0),
    "多渠道开户标志": ("MultiAcctFlag", "数字字符型,限于0-9", 1, 0),
    "对方销售人处投资人基金交易帐号": ("TargetTransactionAccountID", "字符型", 17, 0),
    "投资人收款银行账户户名": ("AcctNameOfInvestorInClearingAgency", "字符型", 200, 0),
    "投资人收款银行账户账号": ("AcctNoOfInvestorInClearingAgency", "字符型", 40, 0),
    "投资人收款银行账户开户行": ("ClearingAgency", "字符型", 20, 0),
    "对帐单寄送方式": ("DeliverWay", "字符型", 8, 0),
    "投资者国家或地区": ("Nationality", "字符型", 3, 0),
    "操作（清算）网点编号": ("NetNo", "字符型", 9, 0),
    "经纪人": ("Broker", "字符型", 12, 0),
    "工作单位名称": ("CorpName", "字符型", 200, 0),
    "证件有效日期": ("CertValidDate", "数字字符型,限于0-9", 8, 0),
    "机构经办人身份证件有效日期": ("InstTranCertValidDate", "数字字符型,限于0-9", 8, 0),
    "机构法人身份证件有效日期": ("InstReprCertValidDate", "数字字符型,限于0-9", 8, 0),
    "客户风险等级": ("ClientRiskRate", "字符型", 1, 0),
    "婚姻状况": ("MarriageStatus", "字符型", 1, 0),
    "家庭人口数": ("FamilyNum", "数值型，其长度不包含小数点，可参与数值计算", 2, 0),
    "家庭资产": ("Penates", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "媒体偏好": ("MediaHobby", "字符型", 1, 0),
    "机构类型": ("InstitutionType", "字符型", 3, 0),
    "投资人英文名": ("EnglishFirstName", "字符型", 20, 0),
    "投资人英文姓": ("EnglishFamliyName", "字符型", 20, 0),
    "行业": ("Vocation", "字符型", 4, 0),
    "企业性质": ("CorpoProperty", "字符型", 2, 0),
    "员工人数": ("StaffNum", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "兴趣爱好类型": ("Hobbytype", "字符型", 2, 0),
    "省/直辖市": ("Province", "字符型", 6, 0),
    "市": ("City", "字符型", 6, 0),
    "县/区": ("County", "字符型", 6, 0),
    "推荐人": ("CommendPerson", "字符型", 40, 0),
    "推荐人类型": ("CommendPersonType", "字符型", 1, 0),
    "受理方式": ("AcceptMethod", "字符型", 1, 0),
    "冻结原因": ("FrozenCause", "数字字符型,限于0-9", 1, 0),
    "冻结截止日期": ("FreezingDeadline", "数字字符型,限于0-9", 8, 0),
    "TA的原确认流水号": ("OriginalSerialNo", "字符型", 20, 0),
    "原申请单编号": ("OriginalAppSheetNo", "字符型", 24, 0),
    "摘要/说明": ("Specification", "字符型", 60, 0),
    "投资者产品代码": ("InvestorProCode", "字符型", 30, 0),
    "辅助身份证明文件类型": ("AuxCertType", "字符型", 3, 0),
    "辅助身份证明文件号码": ("AuxCertCode", "字符型", 40, 0),
    "辅助身份证明文件有效日期": ("AuxCertValidDate", "字符型", 8, 0),
    "IP地址": ("IPAddress", "字符型", 40, 0),
    "MAC地址": ("MACAddress", "字符型", 20, 0),
    "国际移动设备识别码": ("IMEI", "字符型", 20, 0),
    "通用唯一识别码": ("UUID", "字符型", 32, 0),
    "基金代码": ("FundCode", "字符型", 6, 0),
    "巨额赎回处理标志": ("LargeRedemptionFlag", "数字字符型，限于0—9", 1, 0),
    "申请基金份数": ("ApplicationVol", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "申请金额": ("ApplicationAmount", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "销售佣金折扣率": ("DiscountRateOfCommission", "数值型，其长度不包含小数点，可参与数值计算", 5, 4),
    "结算币种": ("CurrencyType", "数字字符型，限于0—9", 3, 0),
    "原申购日期": ("OriginalSubsDate", "数字字符型，限于0—9", 8, 0),
    "交易申请有效天数": ("ValidPeriod", "数值型，其长度不包含小数点，可参与数值计算", 2, 0),
    "预约赎回工作日天数": ("DaysRedemptionInAdvance", "数值型，其长度不包含小数点，可参与数值计算", 5, 0),
    "预约赎回日期": ("RedemptionDateInAdvance", "数字字符型，限于0—9", 8, 0),
    "定期定额申购日期": ("DateOfPeriodicSubs", "数字字符型，限于0—9", 8, 0),
    "TA确认交易流水号": ("TASerialNO", "字符型", 20, 0),
    "定期定额申购期限": ("TermOfPeriodicSubs", "数值型，其长度不包含小数点，可参与数值计算", 5, 0),
    "指定申购日期": ("FutureBuyDate", "数字字符型，限于0—9", 8, 0),
    "对方销售人代码": ("TargetDistributorCode", "字符型", 9, 0),
    "手续费": ("Charge", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "对方网点号": ("TargetBranchCode", "字符型", 9, 0),
    "对方所在地区编号": ("TargetRegionCode", "数字字符型，限于0—9", 4, 0),
    "红利比例": ("DividendRatio", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "转换时的目标基金代码": ("CodeOfTargetFund", "字符型", 6, 0),
    "交易后端收费总额": ("TotalBackendLoad", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "收费类别": ("ShareClass", "字符型", 1, 0),
    "TA的原确认日期": ("OriginalCfmDate", "数字字符型，限于0—9", 8, 0),
    "明细标志": ("DetailFlag", "字符型", 1, 0),
    "原申请日期": ("OriginalAppDate", "数字字符型，限于0—9", 8, 0),
    "默认分红方式": ("DefDividendMethod", "数字字符型，限于0—9", 1, 0),
    "定时定额品种代码": ("VarietyCodeOfPeriodicSubs", "字符型", 5, 0),
    "定时定额申购序号": ("SerialNoOfPeriodicSubs", "字符型", 5, 0),
    "定期定额种类": ("RationType", "字符型", 1, 0),
    "对方基金账号": ("TargetTAAccountID", "字符型", 12, 0),
    "对方TA代码": ("TargetRegistrarCode", "字符型", 18, 0),
    "TA客户编号": ("CustomerNo", "字符型", 12, 0),
    "对方基金份额类别": ("TargetShareType", "字符型", 1, 0),
    "定期定额协议号": ("RationProtocolNo", "字符型", 20, 0),
    "定时定额申购起始日期": ("BeginDateOfPeriodicSubs", "数字字符型，限于0—9", 8, 0),
    "定时定额申购终止日期": ("EndDateOfPeriodicSubs", "数字字符型，限于0—9", 8, 0),
    "定时定额申购每月发送日": ("SendDayOfPeriodicSubs", "数值型，其长度不包含小数点，可参与数值计算", 2, 0),
    "促销活动代码": ("SalesPromotion", "字符型", 3, 0),
    "强制赎回类型": ("ForceRedemptionType", "字符型", 1, 0),
    "带走收益标志": ("TakeIncomeFlag", "字符型", 1, 0),
    "定投目的": ("PurposeOfPeSubs", "字符型", 40, 0),
    "定投频率": ("FrequencyOfPeSubs", "数值型，其长度不包含小数点，可参与数值计算", 5, 0),
    "定投周期单位": ("PeriodSubTimeUnit", "字符型", 1, 0),
    "定投期数": ("BatchNumOfPeSubs", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "资金方式": ("CapitalMode", "字符型", 2, 0),
    "明细资金方式": ("DetailCapticalMode", "字符型", 2, 0),
    "补差费折扣率": ("BackenloadDiscount", "数值型，其长度不包含小数点，可参与数值计算", 5, 4),
    "组合编号": ("CombineNum", "字符型", 6, 0),
    "指定认购日期": ("FutureSubscribeDate", "数字字符型，限于0—9", 8, 0),
    "巨额购买处理标志": ("LargeBuyFlag", "数字字符型，限于0—9", 1, 0),
    "收费类型": ("ChargeType", "字符型", 1, 0),
    "指定费率": ("SpecifyRateFee", "数值型，其长度不包含小数点，可参与数值计算", 9, 8),
    "指定费用": ("SpecifyFee", "数值型，其长度不包含小数点，可参与数值计算", 16, 2),
    "过户原因": ("TransferReason", "字符型", 3, 0),
}


def format_field(value, field_type, field_length, decimal_places=0):
    if pd.isna(value) or not str(value).strip():
        return " " * field_length
    if "数值型" in field_type or "数字字符型" in field_type:
        try:
            cleaned = re.sub(r"[^\d.]", "", str(value))
            num = float(cleaned)
            formatted = f"{num:.{decimal_places}f}".replace(".", "")
            return formatted.zfill(field_length)
        except:
            return " " * field_length
    return str(value).ljust(field_length)[:field_length]

def excel_to_txt(data_file):
    """返回生成的 txt 文件路径"""
    # 解析文件名
    base_name = Path(data_file).stem
    parts = base_name.split('_')
    if len(parts) != 5 or parts[0] != 'OFD':
        raise ValueError("文件名格式必须是：OFD_创建人_接收人_日期_类型.xlsx")
    creator, receiver, date, file_type = parts[1], parts[2], parts[3], parts[4]

    # 读 Excel
    df = pd.read_excel(data_file, dtype=str, keep_default_na=False).fillna("")

    # 生成临时输出文件
    output_file = Path(tempfile.gettempdir()) / f"{base_name}.TXT"
    with open(output_file, 'w', encoding='gb18030', newline='\r\n') as f:
        f.write(f"OFDCFDAT\n22\n{creator}\n{receiver}\n{date}\n00000000\n{file_type}\n\n\n")
        f.write(f"{len(df.columns):08d}\n")
        for col in df.columns:
            f.write(f"{FIELD_MAPPING[col][0]}\n")
        f.write(f"{len(df):016d}\n")
        for _, row in df.iterrows():
            record = []
            for col in df.columns:
                _, field_type, length, decimal = FIELD_MAPPING[col]
                record.append(format_field(row[col], field_type, length, decimal))
            f.write("".join(record) + "\n")
        f.write("OFDCFEND")
    return output_file
# ========================================

# Streamlit 页面
st.set_page_config(page_title="OFD 转换工具", layout="centered")
st.title("📁 OFD Excel → TXT 转换器")
st.markdown("上传符合命名规则的 Excel，一键生成 TXT 并下载。")

uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx"])
if uploaded:
    if st.button("🔧 开始转换"):
        try:
            with st.spinner("正在转换..."):
                # 保存上传文件到临时目录
                temp_excel = Path(tempfile.gettempdir()) / uploaded.name
                with open(temp_excel, "wb") as f:
                    f.write(uploaded.getbuffer())

                # 调用转换函数
                txt_path = excel_to_txt(temp_excel)

            st.success("✅ 转换完成！")
            with open(txt_path, "rb") as f:
                st.download_button(
                    label="⬇️ 下载 TXT",
                    data=f,
                    file_name=txt_path.name,
                    mime="text/plain"
                )

            # 清理临时文件
            os.remove(temp_excel)
            os.remove(txt_path)

        except Exception as e:
            st.error(f"转换失败：{e}")