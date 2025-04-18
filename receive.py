import email
import time
import pickle
import os
import re
from email.header import decode_header
from imapclient import IMAPClient

def save_last_check(last_check, filename='last_check.pkl'):
    """保存上次检查的邮件ID到文件"""
    with open(filename, 'wb') as f:
        pickle.dump(last_check, f)

def load_last_check(filename='last_check.pkl'):
    """从文件加载上次检查的邮件ID"""
    if os.path.exists(filename):
        with open(filename, 'rb') as f:
            return pickle.load(f)
    return set()

def create_imap_client(imap_server, email_account, password):
    """创建并返回一个IMAP客户端连接"""
    print(f"正在连接服务器 {imap_server}...")
    mail = IMAPClient(imap_server, ssl=True, port=993)
    print("正在登录邮箱...")
    mail.login(email_account, password)
    print("正在发送客户端ID信息...")
    mail.id_({"name": "IMAPClient", "version": "2.1.0"})
    print("登录成功！")
    return mail

def parse_email(email_account, password, imap_server="imap.example.com"):
    """
    解析邮件内容的完整函数
    :param email_account: 邮箱账号
    :param password: 邮箱密码/授权码
    :param imap_server: IMAP服务器地址
    :return: 解析后的邮件内容列表
    """
    mail = None
    last_check = load_last_check()  # 从文件加载上次检查的邮件ID
    
    while True:
        try:
            if mail is None:
                mail = create_imap_client(imap_server, email_account, password)
                
                # 列出所有可用的邮箱
                print("正在获取邮箱列表...")
                folders = mail.list_folders()
                print("可用的邮箱:")
                for flags, delimiter, name in folders:
                    print(f"- {name}")
            
            print("开始监听新邮件...")
            
            while True:
                try:
                    # 选择邮箱
                    print("正在选择收件箱...")
                    mail.select_folder('INBOX')
                    print("成功选择收件箱！")
                    
                    # 搜索未读邮件
                    messages = mail.search(['UNSEEN'])
                    current_messages = set(messages)
                    new_messages = current_messages - last_check
                    
                    if new_messages:
                        print(f"\n发现 {len(new_messages)} 封新邮件")
                        
                        for mail_id in new_messages:
                            try:
                                # 获取邮件原始数据
                                msg_data = mail.fetch([mail_id], ['RFC822'])[mail_id]
                                raw_email = msg_data[b'RFC822']
                                
                                # 解析邮件内容
                                msg = email.message_from_bytes(raw_email)
                                
                                # 解析邮件头部信息
                                subject, encoding = decode_header(msg["Subject"])[0]
                                if isinstance(subject, bytes):
                                    subject = subject.decode(encoding or "utf-8")
                                
                                from_ = msg.get("From")
                                date = msg.get("Date")
                                
                                # 初始化邮件内容字典
                                email_content = {
                                    "subject": subject,
                                    "from": from_,
                                    "date": date,
                                    "text": "",
                                    "html": "",
                                    "attachments": []
                                }

                                # 递归解析邮件各部分
                                for part in msg.walk():
                                    content_type = part.get_content_type()
                                    content_disposition = str(part.get("Content-Disposition"))
                                    
                                    try:
                                        # 解析文本内容
                                        if content_type == "text/plain" and "attachment" not in content_disposition:
                                            email_content["text"] = part.get_payload(decode=True).decode(
                                                part.get_content_charset() or "utf-8"
                                            )
                                        # 解析HTML内容
                                        elif content_type == "text/html" and "attachment" not in content_disposition:
                                            email_content["html"] = part.get_payload(decode=True).decode(
                                                part.get_content_charset() or "utf-8"
                                            )
                                        # 处理附件
                                        elif "attachment" in content_disposition:
                                            filename = part.get_filename()
                                            if filename:
                                                filename, encoding = decode_header(filename)[0]
                                                if isinstance(filename, bytes):
                                                    filename = filename.decode(encoding or "utf-8")
                                                
                                                attachment_content = part.get_payload(decode=True)
                                                email_content["attachments"].append({
                                                    "filename": filename,
                                                    "content_type": content_type,
                                                    "size": len(attachment_content),
                                                    "content": attachment_content
                                                })
                                    except Exception as e:
                                        print(f"解析邮件部分时出错: {e}")
                                        continue

                                print(f"\n收到新邮件:")
                                print(f"主题: {email_content['subject']}")
                                print(f"发件人: {email_content['from']}")
                                print(f"日期: {email_content['date']}")
                                print(f"文本内容: {email_content['text'][:200]}...")
                                print(f"HTML内容: {'有' if email_content['html'] else '无'}")
                                print(f"附件数量: {len(email_content['attachments'])}")
                                
                                # 清理文件名，移除特殊字符
                                safe_subject = ''.join(c for c in email_content['subject'] if c.isalnum() or c in (' ', '-', '_')).strip()
                                if not safe_subject:  # 如果清理后文件名为空，使用默认名称
                                    safe_subject = f"email_{int(time.time())}"
                                
                                # 保存HTML内容
                                with open(safe_subject + ".html", "w", encoding='utf-8') as f: 
                                    f.write(email_content['html'])
                                
                                # 提取并保存论文链接
                                if email_content['html']:
                                    paper_links = extract_paper_links(email_content['html'])
                                    if paper_links:
                                        print(f"\n发现 {len(paper_links)} 个论文链接:")
                                        # for link in paper_links:
                                        #     print(f"- https://arxiv.org/abs/{link}")
                                        save_arxiv_links(paper_links)
                                        save_pdf_links(paper_links)
                                
                            except Exception as e:
                                print(f"处理邮件 {mail_id} 时出错: {e}")
                                continue
                        
                        # 更新已检查的邮件ID并保存
                        last_check.update(new_messages)
                        save_last_check(last_check)
                    
                    time.sleep(60)  # 每分钟检查一次新邮件
                    
                except KeyboardInterrupt:
                    print("\n程序已停止")
                    return
                except Exception as e:
                    print(f"发生错误: {e}")
                    print("尝试重新连接...")
                    try:
                        mail.logout()
                    except:
                        pass
                    mail = None
                    time.sleep(5)  # 等待5秒后重试
                    break
        
        except Exception as e:
            print(f"发生严重错误: {e}")
            time.sleep(5)  # 等待5秒后重试
            continue

def extract_paper_links(html_content):
    """从HTML内容中提取论文链接"""
    # 匹配 arXiv 论文链接
    arxiv_pattern = r'https?://arxiv\.org/(?:abs|pdf)/(\d{4}\.\d{4,5}(?:v\d+)?)'
    links = re.findall(arxiv_pattern, html_content)
    return links

def save_arxiv_links(links, filename='arxiv_links.txt'):
    """保存论文链接到文件"""
    with open(filename, 'a', encoding='utf-8') as f:
        for link in links:
            f.write(f"https://arxiv.org/abs/{link}\n")

def save_pdf_links(links, filename='pdf_links.txt'):
    """保存论文链接到文件"""
    with open(filename, 'a', encoding='utf-8') as f:
        for link in links:
            f.write(f"https://arxiv.org/pdf/{link}\n")

# 使用示例
if __name__ == "__main__":
    # 替换为您的邮箱信息
    parse_email(
        email_account="cursor04_lizhi@163.com",
        password="RVQiDbmqK4k6EWB4",  # 注意：这里需要使用授权码，不是邮箱密码
        imap_server="imap.163.com"
    )