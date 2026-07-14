import ssl
import time
import re
import smtplib
import imaplib
import pandas as pd
from email.message import EmailMessage


# ================== 邮箱配置 ==================
SMTP_SERVER = "mail.dlut.edu.cn"
IMAP_SERVER = "mail.dlut.edu.cn"
IMAP_PORT = 993

SENDER = ""
PASSWORD = ""

# 测试用收件人；正式发送时根据 EMAIL_MAP 查找
RECEIVER = "liangzechen2002@gmail.com"
# ============================================


# ================== 姓名-邮箱映射表 ==================
# 来源：0709教育部经费余额梳理-发泽琛(1)-已填充邮箱.xlsx 的"邮箱"sheet
EMAIL_MAP = {
}
# ============================================


# ================== 邮件模板（HTML） ==================
def fmt_balance(v):
    if float(v).is_integer():
        return f"{int(v):,}"
    return f"{v:,.2f}"


def html_to_text(html):
    """把 HTML 正文转成纯文本，用于不支持 HTML 的客户端。"""
    text = html.replace("</p>", "\n\n")
    text = re.sub(r"<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()


TEMPLATES = {
    "间接模版": """<p>{name}老师：</p>
<p>您好！</p>
<p>为切实提高学校哲学社会科学繁荣计划专项资金<strong>（教育部人文社科项目经费）</strong>使用效益，落实财政部、教育部有关工作要求，2024年科学技术研究院人文社科研究院商财务处后，制订了《加强高等学校哲学社会科学繁荣计划专项资金使用管理有关工作举措》（下文简称“举措”）。</p>
<p><strong style="color: #FF0000;">所有间接经费因均已按照经费序时进度上账，</strong>根据《举措》相关要求，您的<strong>项目</strong><strong>间接经费</strong><strong>剩余额度请在</strong><strong style="color: #FF0000;">2026年10月31日前完成经费报销</strong>，<strong style="color: #FF0000;">11月1日后</strong><strong>人文社科研究院将联合财务</strong><strong>处启动</strong><strong>结余</strong><strong style="color: #FF0000;">资金额度收回</strong><strong>。</strong></p>
<p>您的经费账号及余额情况如下：</p>
<p>财务账号为：<span style="background-color: #FFFF00;">{account}</span></p>
<p>经费余额（截至2026.6.30）：<span style="background-color: #FFFF00;">{balance}</span></p>
<p>上述资金额度如存在特殊情况，项目负责人可在<strong>10月31日前</strong>书面向人文社科研究院<strong>提出申请</strong>，明确希望在<strong>2027年批准恢复</strong>的项目结余资金额度，本人签字、经学院盖章知悉同意后，报送人文社科研究院；人文社科研究院将<strong>组织论证</strong>，以确定是否<strong>同意恢复</strong>。</p>
<p>请项目负责人高度重视本次专项资金使用管理工作通知，正确认识个人研究经费的使用事关学校总体的经费额度批复。对以上通知如有不明确的，可随时咨询学院科研院长、科研秘书及人文社科研究院，理工科老师可直接咨询人文社科研究院。</p>
<p>联系人：马小晶</p>
<p>联系电话：84709736</p>""",

    "直接经费（黄色）": """<p>{name}老师：</p>
<p>您好！</p>
<p>为切实提高学校哲学社会科学繁荣计划专项资金<strong>（教育部人文社科项目经费）</strong>使用效益，落实财政部、教育部有关工作要求，2024年科学技术研究院人文社科研究院商财务处后，制订了《加强高等学校哲学社会科学繁荣计划专项资金使用管理有关工作举措》（下文简称“举措”）。</p>
<p>经人文社科研究院核查，您的教育部人文社科项目存在以下问题：1.已经结题（繁荣计划经费结题无法结转统筹金）；2.超过项目研究计划预期2年（含2024年）。根据《举措》相关要求，您的<strong>项目</strong><strong>直接经费</strong><strong>剩余额度请在</strong><strong style="color: #FF0000;">2026年8月31日前完成经费报销</strong>，<strong style="color: #FF0000;">9月1日后</strong><strong>人文社科研究院将联合财务</strong><strong>处启动</strong><strong>结余</strong><strong style="color: #FF0000;">资金额度收回</strong><strong>。</strong></p>
<p>您的经费账号及余额情况如下：</p>
<p>财务账号为：<span style="background-color: #FFFF00;">{account}</span></p>
<p>经费余额（截至2026.6.30）：<span style="background-color: #FFFF00;">{balance}</span></p>
<p>上述资金额度如存在特殊情况，项目负责人可在<strong>8月31日前</strong>书面向人文社科研究院<strong>提出申请</strong>，明确希望在<strong>2027年批准恢复</strong>的项目结余资金额度，本人签字、经学院盖章知悉同意后，报送人文社科研究院；人文社科研究院将<strong>组织论证</strong>，以确定是否<strong>同意恢复</strong>。</p>
<p>请项目负责人高度重视本次专项资金使用管理工作通知，正确认识个人研究经费的使用事关学校总体的经费额度批复。对以上通知如有不明确的，可随时咨询学院科研院长、科研秘书及人文社科研究院，理工科老师可直接咨询人文社科研究院。</p>
<p>联系人：马小晶</p>
<p>联系电话：84709736</p>""",

    "直接经费（橙色）": """<p>{name}老师：</p>
<p>您好！</p>
<p>为切实提高学校哲学社会科学繁荣计划专项资金<strong>（教育部人文社科项目经费）</strong>使用效益，落实财政部、教育部有关工作要求，2024年科学技术研究院人文社科研究院商财务处后，制订了《加强高等学校哲学社会科学繁荣计划专项资金使用管理有关工作举措》（下文简称“举措”）。</p>
<p>经人文社科研究院核查，您的教育部人文社科项目存在以下问题：已超期但未超过研究计划预期2年的在研项目。根据《举措》相关要求，您的<strong>项目</strong><strong>直接经费</strong><strong>剩余额度请在</strong><strong style="color: #FF0000;">2026年10月31日前完成经费报销</strong>。</p>
<p>您的经费账号及余额情况如下：</p>
<p>财务账号为：<span style="background-color: #FFFF00;">{account}</span></p>
<p>经费余额（截至2026.6.30）：<span style="background-color: #FFFF00;">{balance}</span></p>
<p>上述资金额度如无法执行完毕，项目负责人需在<strong>10月31日前</strong>书面向人文社科研究院<strong>提交研究进展报告和经费执行情况</strong>，说明预算执行不到位原因，本人签字、经学院盖章知悉同意后报送人文社科研究院，人文社科研究院将<strong>组织论证</strong>决定是否<strong>收回部分项目结余资金额度</strong>。如存在特殊情况，项目负责人可在<strong>10月31日</strong>前书面向人文社科研究院提出申请，明确希望在2027年批准恢复的项目结余资金额度，本人签字、经学院知悉同意盖章后，人文社科研究院将组织论证，以确定是否同意恢复。</p>
<p>请项目负责人高度重视本次专项资金使用管理工作通知，正确认识个人研究经费的使用事关学校总体的经费额度批复。对以上通知如有不明确的，可随时咨询学院科研院长、科研秘书及人文社科研究院，理工科老师可直接咨询人文社科研究院。</p>
<p>联系人：马小晶</p>
<p>联系电话：84709736</p>""",

    "直接经费（绿色）": """<p>{name}老师：</p>
<p>您好！</p>
<p>为切实提高学校哲学社会科学繁荣计划专项资金<strong>（教育部人文社科项目经费）</strong>使用效益，落实财政部、教育部有关工作要求，2024年科学技术研究院人文社科研究院商财务处后，制订了《加强高等学校哲学社会科学繁荣计划专项资金使用管理有关工作举措》（下文简称“举措”）。</p>
<p>经人文社科研究院核查，您的教育部人文社科项目存在以下问题：未超过研究预期且周期内未达到经费执行序时进度的在研项目。根据《举措》相关要求，您的<strong>项目</strong><strong>直接经费</strong><strong>剩余额度请在</strong><strong style="color: #FF0000;">2026年10月31日前按照序时进度（年初余额+本年拨款90%）完成经费报销</strong>。</p>
<p>您的经费账号及余额情况如下：</p>
<p>财务账号为：<span style="background-color: #FFFF00;">{account}</span></p>
<p>经费余额（截至2026.6.30）：<span style="background-color: #FFFF00;">{balance}</span></p>
<p>上述资金额度如存在特殊情况，项目负责人可在<strong>10月31日前</strong>书面向人文社科研究院<strong>提交研究进展报告和经费执行情况</strong>，说明预算执行不到位原因，本人签字、经学院盖章知悉同意后报送人文社科研究院。</p>
<p>请项目负责人高度重视本次专项资金使用管理工作通知，正确认识个人研究经费的使用事关学校总体的经费额度批复。对以上通知如有不明确的，可随时咨询学院科研院长、科研秘书及人文社科研究院，理工科老师可直接咨询人文社科研究院。</p>
<p>联系人：马小晶</p>
<p>联系电话：84709736</p>""",

    "直接经费（蓝色）": """<p>{name}老师：</p>
<p>您好！</p>
<p>为严格落实学校财务处要求，切实提高哲学社会科学繁荣计划专项资金<strong>（教育部人文社科项目经费）</strong>使用效益，现就加快<strong>教育部人文社科项目经费</strong>执行进度有关事项通知如下。</p>
<p>您的经费账号及余额情况如下：</p>
<p>经费账号：<span style="background-color: #FFFF00;">{account}</span></p>
<p>经费余额（截至2026.6.30）：<span style="background-color: #FFFF00;">{balance}</span></p>
<p>请您尽快统筹安排，在预算内合规支出，切实加快经费执行进度。并于<strong>7月16日前提交一份经费执行方案</strong>（本人签字+学院盖章），送至主楼103室。</p>
<p>方案中应包括项目基本信息、财务信息（含账号、当前余额）以及3个时间节点经费执行情况（2026年11月，2027年6月，2027年12月）。该方案将作为后续执行进度要求的依据，逾期学校将收回经费额度。</p>
<p>联系人：马小晶</p>
<p>联系电话：84709736</p>"""
}
# ============================================


# ================== 发送函数 ==================
def send_by_465_ssl(msg):
    print("正在尝试 465 + SSL 发送邮件...")
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, 465, context=context, timeout=20) as server:
        server.login(SENDER, PASSWORD)
        server.send_message(msg)
    print("465 + SSL 发送成功！")


def send_one(item, max_retries=3):
    """发送单封邮件，失败时自动重试。"""
    for attempt in range(1, max_retries + 1):
        try:
            msg = build_message(item["to_email"], item["subject"], item["text_body"], item["html_body"])
            send_by_465_ssl(msg)
            save_to_sent(msg)
            return True
        except Exception as e:
            print(f"  发送失败（尝试 {attempt}/{max_retries}）：{e}")
            if attempt < max_retries:
                wait = 10 * attempt
                print(f"  {wait} 秒后重试...")
                time.sleep(wait)
    return False


def send_by_587_starttls(msg):
    print("正在尝试 587 + STARTTLS 发送邮件...")
    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVER, 587, timeout=20) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(SENDER, PASSWORD)
        server.send_message(msg)
    print("587 + STARTTLS 发送成功！")


def save_to_sent(msg):
    print("正在保存到已发送...")
    with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as server:
        server.login(SENDER, PASSWORD)

        # 常见“已发送”文件夹名；IMAP 名称含空格时必须加双引号
        candidates = ['Sent', '"Sent Items"', '"Sent Messages"', '已发送', '已发邮件', '发件箱']
        sent_folder = None
        for folder in candidates:
            status, _ = server.select(folder)
            if status == "OK":
                sent_folder = folder
                break

        if sent_folder is None:
            print("常见文件夹名未匹配，当前邮箱文件夹列表：")
            status, folders = server.list()
            for f in folders:
                print(" ", f.decode("utf-8", errors="ignore"))
            raise Exception("未找到'已发送'文件夹")

        server.append(
            sent_folder,
            "",
            imaplib.Time2Internaldate(time.time()),
            msg.as_bytes()
        )
    print(f"已保存到 '{sent_folder}' ！")


def build_message(to_email, subject, text_body, html_body):
    msg = EmailMessage()
    msg["From"] = SENDER
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")
    return msg
# ============================================


# ================== 邮件生成 ==================
def load_recipients(excel_path):
    df = pd.read_excel(excel_path, sheet_name='对应')
    df = df.assign(
        负责人=df['负责人'].astype(str).str.strip(),
        财务编号=df['财务编号'].astype(str).str.strip()
    )
    return df


def generate_emails(excel_path):
    df = load_recipients(excel_path)
    emails = []
    for _, row in df.iterrows():
        name = row['负责人']
        account = row['财务编号']
        balance = fmt_balance(row['项目余额'])
        template_name = str(row['邮件模板']).strip()
        html_body = TEMPLATES[template_name].format(name=name, account=account, balance=balance)
        text_body = html_to_text(html_body)
        subject = "关于加快教育部人文社科项目经费执行进度的通知"
        to_email = EMAIL_MAP.get(name)
        emails.append({
            "name": name,
            "to_email": to_email,
            "subject": subject,
            "text_body": text_body,
            "html_body": html_body,
            "template": template_name,
            "account": account,
            "balance": balance,
        })
    return emails


def preview_to_text(emails, output_path):
    lines = []
    lines.append("=" * 70)
    lines.append("邮件发送清单预览（姓名 + 邮箱 + 正文）")
    lines.append("=" * 70)
    lines.append(f"共 {len(emails)} 封\n")

    for i, item in enumerate(emails, 1):
        lines.append(f"\n{'='*70}")
        lines.append(f"【{i}/{len(emails)}】姓名：{item['name']}  |  模板：{item['template']}")
        lines.append(f"收件邮箱：{item['to_email']}")
        lines.append(f"主题：{item['subject']}")
        lines.append(f"{'='*70}\n")
        lines.append(item['text_body'])
        lines.append("\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"已生成预览文件：{output_path}")
    print(f"共 {len(emails)} 封邮件")
# ============================================


if __name__ == "__main__":
    excel_path = "/Users/xiaochenchener/Documents/Works/科研院/20260710-经费余额邮件/经费余额/发邮件.xlsx"
    preview_path = "/Users/xiaochenchener/Documents/Works/科研院/20260710-经费余额邮件/经费余额/邮件发送清单预览.txt"

    emails = generate_emails(excel_path)
    preview_to_text(emails, preview_path)

    # 当前仅生成预览，不发送。
