# app.py
import streamlit as st
import pandas as pd
import os, re, tempfile
from pathlib import Path
from difflib import get_close_matches

# 常量定义

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



# ==================== 工具函数 ====================
def check_filename(filename):
    """检查文件名是否符合 OFD_创建人_接收人_日期_类型.xlsx 格式"""
    pattern = r"^OFD_.+_.+_\d{8}_.+\.xlsx$"
    return re.match(pattern, filename) is not None


def show_upload_guide():
    """显示上传文件格式提示"""
    with st.expander("ℹ️ 上传文件要求", expanded=True):
        st.markdown("""
        **文件名必须严格遵循以下格式：**  
        `OFD_创建人_接收人_日期_类型.xlsx`  

        例如：  
        ✅ `OFD_002_11_20230801_01.xlsx`  
        ❌ `客户数据.xlsx`  
        ❌ `OFD_002_11.xlsx`
        """)
        st.image("https://via.placeholder.com/600x200?text=Example: OFD_创建人_接收人_日期_类型.xlsx",
                 use_column_width=True)


# ==================== 核心功能 ====================

def format_field(value, field_type, field_length, decimal_places=0):
    """格式化字段（与原代码相同）"""
    if pd.isna(value) or str(value).strip() == "":
        return " " * field_length
    if "数值型" in field_type or "数字字符型" in field_type:
        try:
            cleaned = re.sub(r"[^\d.]", "", str(value))
            num = float(cleaned)
            formatted = f"{num:.{decimal_places}f}".replace(".", "")
            return formatted.zfill(field_length)
        except:
            return " " * field_length
    str_value = str(value).strip()
    try:
        byte_length = len(str_value.encode('gb18030'))
    except UnicodeEncodeError:
        byte_length = len(str_value)
    if byte_length > field_length:
        while byte_length > field_length:
            str_value = str_value[:-1]
            byte_length = len(str_value.encode('gb18030'))
        return str_value
    else:
        return str_value + " " * (field_length - byte_length)


def find_closest_match(column_name, choices, cutoff=0.6):
    """返回最相似的字段名"""
    matches = get_close_matches(column_name, choices, n=1, cutoff=cutoff)
    return matches[0] if matches else None


def interactive_column_mapping(df_columns):
    """交互式列名匹配（仅显示不匹配字段）"""
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}

    # 先处理所有自动匹配
    auto_matched = []
    for col in df_columns:
        if col in FIELD_MAPPING:
            st.session_state.column_mapping[col] = col
            auto_matched.append(col)

    # 仅显示需要处理的列
    unmatched_columns = [col for col in df_columns if col not in auto_matched]
    if not unmatched_columns:
        st.success("✅ 所有列名已自动匹配成功！")
        return st.session_state.column_mapping

    # 显示需要处理的列
    st.subheader("📌 需要确认的列名匹配")
    st.caption(f"发现 {len(unmatched_columns)} 个列需要确认")

    for col in unmatched_columns:
        with st.container(border=True):
            st.markdown(f"**Excel列名:** `{col}`")

            # 尝试模糊匹配
            closest = find_closest_match(col, FIELD_MAPPING.keys())
            if closest:
                col1, col2 = st.columns(2)
                with col1:
                    if st.button(f"👉 接受建议: 「{closest}」",
                                 key=f"accept_{col}",
                                 help=f"匹配到系统字段: {FIELD_MAPPING[closest][0]}"):
                        st.session_state.column_mapping[col] = closest
                        st.rerun()
                with col2:
                    if st.button("🛠 手动选择", key=f"manual_{col}"):
                        st.session_state.current_editing = col
                        st.rerun()
            else:
                st.warning("无自动匹配建议")
                st.session_state.current_editing = col

            # 手动选择模式
            if st.session_state.get('current_editing') == col:
                selected = st.selectbox(
                    "选择对应系统字段:",
                    sorted(FIELD_MAPPING.keys()),
                    key=f"select_{col}"
                )
                if st.button("✅ 确认选择", key=f"confirm_{col}"):
                    st.session_state.column_mapping[col] = selected
                    del st.session_state.current_editing
                    st.rerun()

    # 显示当前进度
    matched_count = len(st.session_state.column_mapping) - len(auto_matched)
    if matched_count > 0:
        st.progress(matched_count / len(unmatched_columns),
                    text=f"匹配进度: {matched_count}/{len(unmatched_columns)}")

    return st.session_state.column_mapping


def excel_to_txt(data_file, column_mapping):
    """支持自定义列名映射的转换函数"""
    df = pd.read_excel(data_file, dtype=str, keep_default_na=False).fillna("")
    df = df.rename(columns=column_mapping)

    # 检查必需字段
    missing_fields = set(FIELD_MAPPING.keys()) - set(df.columns)
    if missing_fields:
        st.warning(f"注意: 缺少建议字段: {', '.join(missing_fields)}")

    # 生成文件名
    base_name = Path(data_file).stem
    parts = base_name.split('_')
    if len(parts) != 5 or parts[0] != 'OFD':
        raise ValueError("文件名格式必须是：OFD_创建人_接收人_日期_类型.xlsx")

    output_file = Path(tempfile.gettempdir()) / f"{base_name}.TXT"
    with open(output_file, 'w', encoding='gb18030', newline='\r\n') as f:
        f.write(f"OFDCFDAT\n22\n{parts[1]}\n{parts[2]}\n{parts[3]}\n00000000\n{parts[4]}\n\n\n")
        f.write(f"{len(df.columns):08d}\n")
        for col in df.columns:
            if col in FIELD_MAPPING:
                f.write(f"{FIELD_MAPPING[col][0]}\n")
        f.write(f"{len(df):016d}\n")
        for _, row in df.iterrows():
            record = []
            for col in df.columns:
                if col in FIELD_MAPPING:
                    _, field_type, length, decimal = FIELD_MAPPING[col]
                    record.append(format_field(row[col], field_type, length, decimal))
            f.write("".join(record) + "\n")
        f.write("OFDCFEND")
    return output_file


# ==================== Streamlit 界面 ====================
st.set_page_config(page_title="OFD 转换工具", layout="centered")
st.title("📁 OFD Excel → TXT 转换器")

# 显示上传指南（始终展示）
show_upload_guide()

# 文件上传区
uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx"],
                            help="请上传符合命名规范的Excel文件")

if uploaded:
    # 实时检查文件名
    if not check_filename(uploaded.name):
        st.error(f"❌ 文件名不符合规范: {uploaded.name}")
        st.markdown("""
        **正确格式示例：**  
        `OFD_销售部_基金公司_20230815_申购.xlsx`
        """)
    else:
        # 显示文件名检查通过
        st.success(f"✅ 文件名验证通过: `{uploaded.name}`")

        # 处理文件内容
        try:
            df = pd.read_excel(uploaded, nrows=0)
            column_mapping = interactive_column_mapping(df.columns.tolist())

            if len(column_mapping) == len(df.columns):
                if st.button("🚀 开始转换", type="primary"):
                    with st.spinner("转换中..."):
                        temp_excel = Path(tempfile.gettempdir()) / uploaded.name
                        with open(temp_excel, "wb") as f:
                            f.write(uploaded.getbuffer())

                        txt_path = excel_to_txt(temp_excel, column_mapping)

                        # 显示转换完成并下载
                        with open(txt_path, "rb") as f:
                            st.download_button(
                                label="⬇️ 下载 TXT 文件",
                                data=f,
                                file_name=txt_path.name,
                                mime="text/plain"
                            )

                        # 清理临时文件
                        os.unlink(temp_excel)
                        os.unlink(txt_path)
        except Exception as e:
            st.error(f"处理错误: {str(e)}")
